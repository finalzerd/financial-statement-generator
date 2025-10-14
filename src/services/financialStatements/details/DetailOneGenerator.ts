import type { TrialBalanceEntry } from '../../../types/financial';
import type { DetailedFinancialData } from '../core/types';
import type { SelectionFirstResult } from '../selection/SelectionFirstClassifier';
import type { IAccountMappingProvider } from '../mapping/IAccountMappingProvider';
import { AccountMappingUtils } from '../../../types/accountMapping';
import { FinancialCalculations } from '../../financialCalculations';
import type { DetailOneMode } from '../../../types/detailSettings';

export class DetailOneGenerator {
  /**
   * Detail One: Cost of Goods Sold / Cost of Services
   * Handles both inventory-based and service-based businesses
   * Uses global data for optimized calculations
   */
  static generateDetailOne(
    trialBalanceData: TrialBalanceEntry[], 
    globalData: DetailedFinancialData,
    selection?: SelectionFirstResult,
    mode: DetailOneMode = 'auto',
    provider?: IAccountMappingProvider,
    startingRow: number = 1
  ): any[][] {
    const detailNotes: any[][] = [];
    let currentRow = startingRow;
    const pushRow = (row: any[]): number => {
      detailNotes.push(row);
      const insertedRow = currentRow;
      currentRow += 1;
      return insertedRow;
    };
    const detectedInventory = globalData.flags?.hasInventory ?? FinancialCalculations.checkHasInventory(trialBalanceData);
    const effectiveMode: DetailOneMode = mode === 'auto'
      ? (detectedInventory ? 'inventory' : 'service')
      : mode;
    
    // Header
    pushRow(['รายละเอียดประกอบที่ 1', '', '', '', '', '', '', '', 'หน่วย:บาท']);
    pushRow(['', '', '', '', '', '', '', '', '']);
    
    if (effectiveMode === 'both') {
      const hasServiceRows = this.addServiceBusinessDetail(pushRow, trialBalanceData, provider, selection);
      if (hasServiceRows) {
        pushRow(['', '', '', '', '', '', '', '', '']);
      }
      this.addInventoryBusinessDetail(pushRow, trialBalanceData, globalData, selection);
      return detailNotes;
    }

    if (effectiveMode === 'service') {
      this.addServiceBusinessDetail(pushRow, trialBalanceData, provider, selection);
      return detailNotes;
    }

    // Default to inventory calculation
    this.addInventoryBusinessDetail(pushRow, trialBalanceData, globalData, selection);
    
    return detailNotes;
  }

  /**
   * Add service business cost details
   */
  private static addServiceBusinessDetail(
    pushRow: (row: any[]) => number,
    trialBalanceData: TrialBalanceEntry[],
    provider?: IAccountMappingProvider,
    selection?: SelectionFirstResult
  ): boolean {
    // Section header
    pushRow(['', 'ต้นทุนการให้บริการ', '', '', '', '', '', '', '']);

    const serviceAccounts = this.getServiceCostAccounts(trialBalanceData, provider, selection);

    if (serviceAccounts.length === 0) {
      pushRow(['', '', 'ไม่พบบัญชีที่ตรงกับเกณฑ์สำหรับต้นทุนการให้บริการ', '', '', '', '', '', '']);
      pushRow(['', '', 'รวมต้นทุนการให้บริการ', '', '', '', '', '', 0]);
      return false;
    }

    // Track first and last account rows explicitly
    let firstDetailRow = 0;
    let firstRowCaptured = false;
    let lastDetailRow = 0;
    for (const account of serviceAccounts) {
      const rowIndex = pushRow([
        '',
        '',
        account.accountName,
        '',
        '',
        '',
        '',
        '',
        account.current
      ]);
      if (!firstRowCaptured) { firstRowCaptured = true; firstDetailRow = rowIndex; }
      lastDetailRow = rowIndex;
    }

    const totalFormula = lastDetailRow > 0
      ? { f: `SUM(I${firstDetailRow}:I${lastDetailRow})` }
      : 0;
    pushRow([
      '',
      '',
      'รวมต้นทุนการให้บริการ',
      '',
      '',
      '',
      '',
      '',
      totalFormula
    ]);

    return true;
  }

  /**
   * Add inventory business cost of goods sold details
   */
  private static addInventoryBusinessDetail(
    pushRow: (row: any[]) => number,
    trialBalanceData: TrialBalanceEntry[], 
    globalData: DetailedFinancialData,
    selection?: SelectionFirstResult
  ): void {
    pushRow(['ต้นทุนสินค้าที่ขาย', '', '', '', '', '', '', '', '']);
    
    // Prefer selection-first inventory totals with legacy fallback to global data
    const inventoryTotals = selection?.totals?.inventory;
    const currentInventory = inventoryTotals?.current ?? globalData.noteCalculations.inventory.total.current;
    const previousInventory = inventoryTotals?.previous ?? globalData.noteCalculations.inventory.total.previous;
    
    pushRow(['', 'สินค้าคงเหลือต้นงวด', '', '', '', '', '', '', previousInventory]);
    
    // Process purchases (still need individual account details)
    let totalPurchases = 0;
    const purchaseAmount = this.resolveSelectionAmount(selection, '5010', () => Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5010, 5010)));
    if (purchaseAmount > 0) {
      pushRow(['', 'บวก', 'ซื้อสินค้า', '', '', '', '', '', purchaseAmount]);
      totalPurchases += purchaseAmount;
    }
    
    // Purchase returns - still need individual calculation
    const returnAmount = this.resolveSelectionAmount(selection, '5010.1', () => Math.abs(this.getAccountBalance(trialBalanceData, ['5010.1'])));
    if (returnAmount > 0) {
      pushRow(['', 'หัก', 'ส่งคืนสินค้า', '', '', '', '', '', returnAmount]);
      totalPurchases -= returnAmount;
    }
    
    const availableForSale = previousInventory + totalPurchases;
    pushRow(['', '', 'สินค้าที่มีไว้เพื่อขาย', '', '', '', '', '', availableForSale]);
    pushRow(['', 'หัก', 'สินค้าคงเหลือปลายงวด', '', '', '', '', '', currentInventory]);
    
    const costOfGoodsSold = availableForSale - currentInventory;
    pushRow(['', '', 'ต้นทุนสินค้าที่ขาย', '', '', '', '', '', costOfGoodsSold]);
  }

  /**
   * Helper method: Sum accounts by numeric range
   */
  private static sumAccountsByNumericRange(
    trialBalanceData: TrialBalanceEntry[], 
    startCode: number, 
    endCode: number
  ): number {
    return trialBalanceData
      .filter(entry => {
        const code = parseInt(entry.accountCode || '0');
        return code >= startCode && code <= endCode;
      })
      .reduce((sum, entry) => sum + (entry.balance || 0), 0);
  }

  /**
   * Helper method: Get account balance by specific codes
   */
  private static getAccountBalance(
    trialBalanceData: TrialBalanceEntry[], 
    accountCodes: string[]
  ): number {
    return trialBalanceData
      .filter(entry => accountCodes.includes(entry.accountCode || ''))
      .reduce((sum, entry) => sum + (entry.balance || 0), 0);
  }

  private static getServiceCostAccounts(
    trialBalanceData: TrialBalanceEntry[],
    provider?: IAccountMappingProvider,
    selection?: SelectionFirstResult
  ): Array<{ accountCode: string; accountName: string; current: number; previous: number }> {
    const selectionAccounts = selection?.byCategory?.detail_service_costs ?? [];

    if (selectionAccounts.length > 0) {
      const normalizedSelection = selectionAccounts
        .map(account => ({
          accountCode: account.accountCode,
          accountName: account.accountName,
          current: Math.abs(account.rawCurrent ?? account.current ?? 0),
          previous: Math.abs(account.rawPrevious ?? account.previous ?? 0)
        }))
        .filter(account => account.current !== 0 || account.previous !== 0);

      if (normalizedSelection.length > 0) {
        return this.sortAccounts(normalizedSelection);
      }
    }

    const rules = provider?.getRules('detail_service_costs');
    const rawEntries = rules
      ? AccountMappingUtils.getMatchingAccounts(trialBalanceData, rules)
      : trialBalanceData.filter(entry => {
          const code = parseFloat(entry.accountCode || '0');
          return code >= 5000 && code <= 5099;
        });

    const normalized = rawEntries
      .map(entry => {
        const currentAmount = Math.abs((entry.balance ?? entry.currentBalance ?? 0) as number);
        const previousAmount = Math.abs((entry.previousBalance ?? 0) as number);
        return {
          accountCode: entry.accountCode || '',
          accountName: entry.accountName || `บัญชี ${entry.accountCode}`,
          current: currentAmount,
          previous: previousAmount
        };
      })
      .filter(account => account.current !== 0 || account.previous !== 0);

    return this.sortAccounts(normalized);
  }

  private static sortAccounts(
    accounts: Array<{ accountCode: string; accountName: string; current: number; previous: number }>
  ): Array<{ accountCode: string; accountName: string; current: number; previous: number }> {
    return accounts.sort((a, b) => {
      const aCode = a.accountCode || '';
      const bCode = b.accountCode || '';
      const aNum = parseFloat(aCode);
      const bNum = parseFloat(bCode);
      if (Number.isFinite(aNum) && Number.isFinite(bNum)) {
        return aNum - bNum;
      }
      return aCode.localeCompare(bCode);
    });
  }

  private static resolveSelectionAmount(
    selection: SelectionFirstResult | undefined,
    accountCode: string,
    fallback: () => number
  ): number {
    const normalizedCode = accountCode.trim();
    const account = selection?.byAccount?.[normalizedCode];
    if (!account) {
      return fallback();
    }
    const rawValue = account.rawCurrent ?? account.current ?? 0;
    const amount = Math.abs(rawValue);
    if (amount === 0) {
      return fallback();
    }
    return amount;
  }
}
