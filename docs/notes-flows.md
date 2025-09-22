# Notes Function Flows

Function-level flows for each note generator that uses the row tracking architecture. Labels are sanitized for Mermaid.

## Conventions
- Nodes avoid quotes, parentheses, and special punctuation
- Multi-year handling shown when relevant
- Tracker fields: headerRows, yearHeaderRows, detailRows, totalRows, unitRows

---

## Cash Note (Note 7)

```mermaid
flowchart TD
  A[CashNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Read totals from globalData.noteCalculations.cash]
  C --> D[Guard return if all totals are zero]
  D --> E[Push header row and unit row]
  E --> F[Push year header based on processingType]
  F --> G[If cashAmount present push เงินสดในมือ]
  G --> H[If bankAmount present push เงินฝากธนาคาร]
  H --> I[Push total row รวม]
  I --> J[Push spacer row]
  J --> K[Return tracker]
```

## Trade Receivables Note (Note 8)

```mermaid
flowchart TD
  A[TradeReceivablesNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Get receivable accounts from globalData.individualAccounts.receivables]
  C --> D[Read totals from globalData.noteCalculations.receivables]
  D --> E[Guard return if totals and accounts empty]
  E --> F[Push header and unit rows]
  F --> G[Push year header]
  G --> H[For each receivable account push detail row]
  H --> I[Push total row รวม]
  I --> J[Push spacer]
  J --> K[Return tracker]
```

## Trade Payables Note (Note 12)

```mermaid
flowchart TD
  A[TradePayablesNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Get payable accounts from globalData.individualAccounts.payables]
  C --> D[Read totals from globalData.noteCalculations.payables]
  D --> E[Guard return if totals and accounts empty]
  E --> F[Push header and unit rows]
  F --> G[Push year header]
  G --> H[For each payable account push detail row]
  H --> I[Push total row รวม]
  I --> J[Push spacer]
  J --> K[Return tracker]
```

## PPE Note (Note 10)

```mermaid
flowchart TD
  A[PPENoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Filter asset accounts 1610-1659]
  C --> D[Filter depreciation accounts 1610-1659 dot accounts]
  D --> E[Guard return if no balances]
  E --> F[Push header and unit rows]
  F --> G[Push column headers based on processingType]
  G --> H[Push section header ราคาทุนเดิม]
  H --> I[Loop asset accounts compute purchases disposals and push detail]
  I --> J[Push asset totals with SUM formulas]
  J --> K[Push section header ค่าเสื่อมราคาสะสม]
  K --> L[Loop depreciation accounts compute expense disposal and push detail]
  L --> M[Push depreciation totals with SUM formulas]
  M --> N[Push net book value formulas]
  N --> O[If expense present push ค่าเสื่อมราคา summary]
  O --> P[Push spacer]
  P --> Q[Return tracker]
```

## Short Term Loans Note (Note 5)

```mermaid
flowchart TD
  A[ShortTermLoansNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Sum accounts 1141]
  C --> D[Sum previous balances if multi year]
  D --> E[Guard return if both zero]
  E --> F[Push header and unit rows]
  F --> G[Push year header]
  G --> H[Push detail เงินให้กู้ยืมระยะสั้น]
  H --> I[Push total as formula referencing detail]
  I --> J[Push spacer]
  J --> K[Return tracker]
```

## Other Assets Note

```mermaid
flowchart TD
  A[OtherAssetsNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Sum accounts 1660-1700]
  C --> D[Sum previous balances if multi year]
  D --> E[Guard return if both zero]
  E --> F[Push header and unit rows]
  F --> G[Push year header]
  G --> H[Push detail สินทรัพย์อื่น]
  H --> I[Push total as formula referencing detail]
  I --> J[Push spacer]
  J --> K[Return tracker]
```

## Long Term Loans from FI Note

```mermaid
flowchart TD
  A[LongTermLoansNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Sum 2120-2123 minus 2121]
  C --> D[Sum previous balances if multi year]
  D --> E[Guard return if both zero]
  E --> F[Push header and unit rows]
  F --> G[Push year header]
  G --> H[Push main detail เงินกู้ยืมระยะยาวจากสถาบันการเงิน]
  H --> I[Push total as formula referencing main detail]
  I --> J[Push current portion 10 percent placeholder]
  J --> K[Push net long term loans]
  K --> L[Push spacer]
  L --> M[Return tracker]
```

## Other Long Term Loans Note

```mermaid
flowchart TD
  A[OtherLongTermLoansNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Sum accounts 2050-2052]
  C --> D[Sum previous balances if multi year]
  D --> E[Guard return if both zero]
  E --> F[Push header and unit rows]
  F --> G[Push year header]
  G --> H[Push detail เงินกู้ยืมระยะยาว]
  H --> I[Push total as formula referencing detail]
  I --> J[Push spacer]
  J --> K[Return tracker]
```

## Related Party Loans Note

```mermaid
flowchart TD
  A[RelatedPartyLoansNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Sum account 2100]
  C --> D[Sum previous balances if multi year]
  D --> E[Guard return if both zero]
  E --> F[Push header and unit rows]
  F --> G[Push year header]
  G --> H[Push detail เงินกู้ยืมระยะยาวจากบุคคลที่เกี่ยวข้อง]
  H --> I[Push total as formula referencing detail]
  I --> J[Push spacer]
  J --> K[Return tracker]
```

## Other Income Note

```mermaid
flowchart TD
  A[OtherIncomeNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Sum accounts 4110-4999 using P and L rule]
  C --> D[Sum previous balances 4110-4999]
  D --> E[Guard return if both zero]
  E --> F[Push header and unit rows]
  F --> G[Push year header]
  G --> H[Loop other income accounts and push detail]
  H --> I[If multiple items push total รวม]
  I --> J[Push spacer]
  J --> K[Return tracker]
```

## Expenses By Nature Note

```mermaid
flowchart TD
  A[ExpensesByNatureNoteGenerator.generateWithRowTracking] --> B[Init NoteRowTracker]
  B --> C[Push header and unit rows]
  C --> D[Push year header]
  D --> E[Push template categories]
  E --> F[Push total row]
  F --> G[Push spacer]
  G --> H[Return tracker]
```
