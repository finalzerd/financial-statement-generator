# Balance Sheet Flow Diagrams

This page documents how the Balance Sheet’s Assets and Liabilities/Equity are constructed in this project.

Tip: You can view these Mermaid diagrams directly in VS Code (Markdown preview) or on GitHub.

## High-level build flow

```mermaid
flowchart TD
  A[Start: Trial Balance + Company Info] --> B[GlobalDataExtractor.extract]
  B --> B1[Compute foundation totals\n- Cash 1000-1099\n- Receivables 1140-1215\n- Payables 2010-2999 excl\n- Inventory 1500-1519\n- PPE and other aggregates]
  B --> B2[Discover individual accounts\n- cash receivables payables]
  B --> B3[Flags and context\n- Inventory vs service\n- Company type\n- Multi-year]

  B --> C[AssetsBuilder.build]
  C --> C1[Current Assets\n- Cash, Receivables, Inventory, Prepaids]
  C1 --> C2[Non-current Assets\n- PPE net, Other assets]
  C2 --> C3[Totals via formulas\n- Subtotals\n- Total Assets]
  C3 --> E[BS_Assets worksheet]

  B --> D[LiabilitiesBuilder.build]
  D --> D1[Current Liabilities\n- Trade and Other payables\n- Short-term items]
  D1 --> D2[Non-current Liabilities\n- Long-term loans]
  D2 --> D3[Equity\n- Capital, Legal reserve, Retained earnings]
  D3 --> D4[Totals via formulas\n- Total Liabilities plus Equity]
  D4 --> F[BS_Liabilities worksheet]

  E --> G[ExcelJSFormatter formatting]
  F --> G
  G --> H[Validation\nAssets equals Liabilities plus Equity]
  H --> I[Downloadable Excel]
```

## Function call flow (exact methods)

```mermaid
flowchart TD
  A[FinancialStatementGenerator.generateFinancialStatements] --> B[GlobalDataExtractor.extract]
  A --> C[AssetsBuilder.build]
  A --> D[LiabilitiesBuilder.build]
  A --> E[downloadAsExcel]
  E --> F[ExcelJSFormatter.addDataToWorksheet BS_Assets]
  E --> G[ExcelJSFormatter.addDataToWorksheet BS_Liabilities]
  E --> H[ExcelJSFormatter.formatBalanceSheetAssets]

  %% Assets
  C --> C1[FinancialCalculations.sumAccountsByNumericRange]
  C --> C2[FinancialCalculations.sumPreviousBalanceByNumericRange]
  C --> C3[FinancialCalculations.buildSumFormula]

  %% Liabilities and Equity
  D --> D1[buildCurrentLiabilitiesSection]
  D --> D2[buildNonCurrentLiabilitiesSection]
  D --> D3[buildEquitySection]

  D1 --> DC[sumAccountsByNumericRange and sumPrevious]
  D1 --> DF[buildSumFormula]
  D2 --> DG[sumAccountsByNumericRange and sumPrevious]
  D2 --> DH[buildSumFormula]
  D3 --> DI[getSingleAccountBalance 3010]
  D3 --> DJ[getOpeningRetainedEarnings]
  D3 --> DK[calculateCurrentYearProfit]
  D3 --> DL[buildSumFormula]
```

## How to view

- VS Code: open this file and use Markdown Preview (Ctrl+Shift+V). If Mermaid doesn’t render, install the extension “Markdown Preview Mermaid Support” (bierner.markdown-mermaid).
- GitHub: commit and push; open this file on GitHub and Mermaid will render automatically.
