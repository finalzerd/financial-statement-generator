import type { TrialBalanceEntry } from '../../../types/financial';
import type { DetailedFinancialData } from '../core/types';
import type { SelectionFirstResult } from '../selection/SelectionFirstClassifier';
import { FinancialCalculations } from '../../financialCalculations';

export class DetailOneGenerator {
  /**
   * Detail One: Cost of Goods Sold / Cost of Services
   * Handles both inventory-based and service-based businesses
   * Uses global data for optimized calculations
   */
  static generateDetailOne(
    trialBalanceData: TrialBalanceEntry[], 
    globalData: DetailedFinancialData,
    selection?: SelectionFirstResult
  ): any[][] {
    const detailNotes: any[][] = [];
    const hasInventory = globalData.flags?.hasInventory ?? FinancialCalculations.checkHasInventory(trialBalanceData);
    
    // Header
    detailNotes.push(['รายละเอียดประกอบที่ 1', '', '', '', '', '', '', '', 'หน่วย:บาท']);
    detailNotes.push(['', '', '', '', '', '', '', '', '']);
    
    if (!hasInventory) {
      // Service business - use global data for cost of services
      this.addServiceBusinessDetail(detailNotes);
    } else {
      // Inventory-based business - use optimized global data
      this.addInventoryBusinessDetail(detailNotes, trialBalanceData, globalData, selection);
    }
    
    return detailNotes;
  }

  /**
   * Add service business cost details
   */
  private static addServiceBusinessDetail(detailNotes: any[][]): void {
    detailNotes.push(['ต้นทุนการให้บริการ', '', '', '', '', '', '', '', '']);
    detailNotes.push(['', 'ค่าใช้จ่ายอื่นๆ ในการให้บริการ', '', '', '', '', '', '', '...']);
    detailNotes.push(['', 'รวม', '', '', '', '', '', '', { f: 'I' + (detailNotes.length) }]);
  }

  /**
   * Add inventory business cost of goods sold details
   */
  private static addInventoryBusinessDetail(
    detailNotes: any[][],
    trialBalanceData: TrialBalanceEntry[], 
    globalData: DetailedFinancialData,
    selection?: SelectionFirstResult
  ): void {
    detailNotes.push(['ต้นทุนสินค้าที่ขาย', '', '', '', '', '', '', '', '']);
    
    // Prefer selection-first inventory totals with legacy fallback to global data
    const inventoryTotals = selection?.totals?.inventory;
    const currentInventory = inventoryTotals?.current ?? globalData.noteCalculations.inventory.total.current;
    const previousInventory = inventoryTotals?.previous ?? globalData.noteCalculations.inventory.total.previous;
    
    detailNotes.push(['', 'สินค้าคงเหลือต้นงวด', '', '', '', '', '', '', previousInventory]);
    
    // Process purchases (still need individual account details)
    let totalPurchases = 0;
    const purchaseAmount = this.resolveSelectionAmount(selection, '5010', () => Math.abs(this.sumAccountsByNumericRange(trialBalanceData, 5010, 5010)));
    if (purchaseAmount > 0) {
      detailNotes.push(['', 'บวก', 'ซื้อสินค้า', '', '', '', '', '', purchaseAmount]);
      totalPurchases += purchaseAmount;
    }
    
    // Purchase returns - still need individual calculation
    const returnAmount = this.resolveSelectionAmount(selection, '5010.1', () => Math.abs(this.getAccountBalance(trialBalanceData, ['5010.1'])));
    if (returnAmount > 0) {
      detailNotes.push(['', 'หัก', 'ส่งคืนสินค้า', '', '', '', '', '', returnAmount]);
      totalPurchases -= returnAmount;
    }
    
    const availableForSale = previousInventory + totalPurchases;
    detailNotes.push(['', '', 'สินค้าที่มีไว้เพื่อขาย', '', '', '', '', '', availableForSale]);
    detailNotes.push(['', 'หัก', 'สินค้าคงเหลือปลายงวด', '', '', '', '', '', currentInventory]);
    
    const costOfGoodsSold = availableForSale - currentInventory;
    detailNotes.push(['', '', 'ต้นทุนสินค้าที่ขาย', '', '', '', '', '', costOfGoodsSold]);
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
