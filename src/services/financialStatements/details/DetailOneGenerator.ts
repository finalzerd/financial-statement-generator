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
      this.addInventoryBusinessDetail(pushRow, trialBalanceData, globalData, selection, provider);
      return detailNotes;
    }

    if (effectiveMode === 'service') {
      this.addServiceBusinessDetail(pushRow, trialBalanceData, provider, selection);
      return detailNotes;
    }

  // Default to inventory calculation
  this.addInventoryBusinessDetail(pushRow, trialBalanceData, globalData, selection, provider);
    
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
    selection?: SelectionFirstResult,
    provider?: IAccountMappingProvider
  ): void {
    pushRow(['ต้นทุนสินค้าที่ขาย', '', '', '', '', '', '', '', '']);
    
    // Prefer selection-first inventory totals with legacy fallback to global data
    const inventoryTotals = selection?.totals?.inventory;
    const currentInventory = inventoryTotals?.current ?? globalData.noteCalculations.inventory.total.current;
    const previousInventory = inventoryTotals?.previous ?? globalData.noteCalculations.inventory.total.previous;
    
    pushRow(['', 'สินค้าคงเหลือต้นงวด', '', '', '', '', '', '', previousInventory]);
    
    // Process purchases, returns, discounts as grouped lines using selection-first (with provider/legacy fallbacks)
    const purchases = this.sumInventoryCategory(trialBalanceData, provider, selection, 'inventory_purchases');
    const purchaseReturns = this.sumInventoryCategory(trialBalanceData, provider, selection, 'inventory_purchase_returns');
    const purchaseDiscounts = this.sumInventoryCategory(trialBalanceData, provider, selection, 'inventory_purchase_discounts');

    if (purchases.current > 0) {
      pushRow(['', 'บวก', 'ซื้อสินค้า', '', '', '', '', '', purchases.current]);
    }
    if (purchaseReturns.current > 0) {
      pushRow(['', 'หัก', 'ส่งคืนสินค้า', '', '', '', '', '', purchaseReturns.current]);
    }
    if (purchaseDiscounts.current > 0) {
      pushRow(['', 'หัก', 'ส่วนลดรับ', '', '', '', '', '', purchaseDiscounts.current]);
    }
    
  const availableForSale = previousInventory + purchases.current - purchaseReturns.current - purchaseDiscounts.current;
    pushRow(['', '', 'สินค้าที่มีไว้เพื่อขาย', '', '', '', '', '', availableForSale]);
    pushRow(['', 'หัก', 'สินค้าคงเหลือปลายงวด', '', '', '', '', '', currentInventory]);
    
    const costOfGoodsSold = availableForSale - currentInventory;
    pushRow(['', '', 'ต้นทุนสินค้าที่ขาย', '', '', '', '', '', costOfGoodsSold]);
  }

  // (legacy helpers removed: sumAccountsByNumericRange, getAccountBalance)

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

  // (legacy helper removed: getInventoryPurchaseAccounts)

  // NEW: Sum totals for inventory purchases/returns/discounts categories
  private static sumInventoryCategory(
    trialBalanceData: TrialBalanceEntry[],
    provider: IAccountMappingProvider | undefined,
    selection: SelectionFirstResult | undefined,
    category: 'inventory_purchases' | 'inventory_purchase_returns' | 'inventory_purchase_discounts'
  ): { current: number; previous: number } {
    // Prefer selection-first
    const selectionAccounts = (selection?.byCategory as any)?.[category] ?? [];
    if (selectionAccounts.length > 0) {
      return selectionAccounts.reduce((acc: { current: number; previous: number }, a: any) => {
        acc.current += Math.abs(a.rawCurrent ?? a.current ?? 0);
        acc.previous += Math.abs(a.rawPrevious ?? a.previous ?? 0);
        return acc;
      }, { current: 0, previous: 0 });
    }

    // Provider rules fallback
    const rules = provider?.getRules(category);
    const rawEntries = rules
      ? AccountMappingUtils.getMatchingAccounts(trialBalanceData as any[], rules)
      : trialBalanceData.filter(entry => {
          const codeStr = (entry.accountCode || '').trim();
          const codeNum = parseFloat(codeStr || '0');
          const normalized = codeStr.replace(/\s+/g, '');
          if (category === 'inventory_purchases') return codeNum === 5010;
          if (category === 'inventory_purchase_returns') return normalized === '5010.1' || codeNum === 5010.1;
          return normalized === '5010.2' || codeNum === 5010.2; // discounts
        });

    return rawEntries.reduce((acc, e) => {
      acc.current += Math.abs((e.balance ?? (e as any).currentBalance ?? 0) as number);
      acc.previous += Math.abs((e.previousBalance ?? 0) as number);
      return acc;
    }, { current: 0, previous: 0 });
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

  // (legacy helper removed: resolveSelectionAmount)
}
