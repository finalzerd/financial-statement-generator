# PPE Movement Note - VBA Structure Implementation

I've successfully recreated the PPE (Property, Plant & Equipment) note to follow the VBA code structure as requested. Here's what was implemented:

## Key Changes Made

### 1. **Enhanced Data Structure**
```typescript
// Updated TrialBalanceEntry to preserve original CSV data
export interface TrialBalanceEntry {
  accountCode: string;
  accountName: string;
  debitAmount: number;
  creditAmount: number;
  balance: number;
  previousBalance?: number;  // ยอดยกมาต้นงวด from CSV
  currentBalance?: number;   // ยอดยกมางวดนี้ from CSV
}
```

### 2. **Updated CSV Processor**
```typescript
// Now preserves original CSV data for movement calculations
static convertToCurrentYear(csvEntries: CSVTrialBalanceEntry[]): TrialBalanceEntry[] {
  return csvEntries.map(entry => ({
    accountCode: entry.รหัสบัญชี,
    accountName: entry.ชื่อบัญชี,
    balance: entry.เดบิต ? entry.เดบิต : -entry.เครดิต,
    debitAmount: entry.เดบิต || 0,
    creditAmount: entry.เครดิต || 0,
    // Preserve original CSV data for PPE movement calculation
    previousBalance: entry.ยอดยกมาต้นงวด,
    currentBalance: entry.ยอดยกมางวดนี้
  }));
}
```

### 3. **New PPE Movement Note Structure**

The PPE note now follows VBA format with these columns:

| Column | Thai Header | Description |
|--------|-------------|-------------|
| 1 | ยอดยกมา ณ วันที่ 1 มกราคม | Previous amount = `ยอดยกมาต้นงวด` |
| 2 | เพิ่มขึ้น ระหว่างงวด | Additions during period |
| 3 | ลดลง ระหว่างงวด | Disposals during period |
| 4 | ยอดสิ้นงวด ณ วันที่ 31 ธันวาคม | Current amount = Previous + Current |

## Calculation Logic

### **As Per Your Requirements:**

1. **Previous Amount** = `ยอดยกมาต้นงวด` (Previous balance from CSV)
2. **Current Amount** = `ยอดยกมาต้นงวด + ยอดยกมางวดนี้` (Previous + Current balance from CSV)
3. **Movement Calculation:**
   - If Current > Previous → **Addition** = Current - Previous
   - If Current < Previous → **Disposal** = Previous - Current

### **Code Implementation:**
```typescript
// Calculate based on CSV structure as per user requirements
for (const entry of matchingEntries) {
  const prevBalance = Math.abs(entry.previousBalance || 0);
  const currBalance = Math.abs(entry.currentBalance || 0);
  
  previousAmount += prevBalance;
  currentAmount += prevBalance + currBalance;  // Previous + Current
}

// Calculate movements
const netMovement = currentAmount - previousAmount;
const additions = netMovement > 0 ? netMovement : 0;
const disposals = netMovement < 0 ? Math.abs(netMovement) : 0;
```

## Excel Output Structure

The generated Excel file will now contain a PPE note that looks like this:

```
7    ที่ดิน อาคารและอุปกรณ์
                    ยอดยกมา        เพิ่มขึ้น       ลดลง         ยอดสิ้นงวด
                    ณ วันที่ 1 มกราคม  ระหว่างงวด     ระหว่างงวด     ณ วันที่ 31 ธันวาคม
                    2024           2024         2024         2024
                    บาท            บาท          บาท          บาท
     ที่ดิน           100,000.00     50,000.00    -           150,000.00
     อาคารและส่วนปรับปรุง  200,000.00     -           10,000.00    190,000.00
     เครื่องจักรและอุปกรณ์  150,000.00     25,000.00    -           175,000.00
     รวม              450,000.00     75,000.00    10,000.00    515,000.00
```

## PPE Account Categories

The system tracks these PPE categories:
- **ที่ดิน** (Land): Account code `1610`
- **อาคารและส่วนปรับปรุง** (Buildings & Improvements): Account code `1615`
- **เครื่องจักรและอุปกรณ์** (Machinery & Equipment): Account codes `1640`, `1641`, `1645`

## Benefits

1. **VBA Compatibility**: Matches the original VBA output format exactly
2. **Accurate Movement Tracking**: Uses actual CSV data for precise calculations
3. **Automatic Addition/Disposal Detection**: Based on net movement direction
4. **Thai Accounting Standards**: Follows proper Thai financial reporting format
5. **Debug Support**: Includes console logging for troubleshooting

## Usage

The PPE note will automatically appear in the generated Excel file when:
1. PPE accounts (1610, 1615, 1640, 1641, 1645) have balances > 0
2. CSV data contains both previous and current balance information
3. Financial statements are generated with the updated system

This implementation ensures that the Excel output closely matches the VBA code structure while providing accurate movement analysis based on the actual trial balance data.
