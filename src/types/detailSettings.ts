export type DetailOneMode = 'auto' | 'service' | 'inventory' | 'both';

export interface DetailSettings {
  detailOneMode: DetailOneMode;
  updatedAt?: string | null;
  /** Indicates whether the preference was persisted on the backend */
  persisted?: boolean;
}

export const DETAIL_ONE_MODE_OPTIONS: Array<{ value: DetailOneMode; label: string; description: string }> = [
  { value: 'auto', label: 'โหมดอัตโนมัติ', description: 'เลือกตามการตรวจพบสินค้าคงเหลือ (ค่าเริ่มต้น)' },
  { value: 'service', label: 'ธุรกิจให้บริการ', description: 'ใช้รูปแบบต้นทุนการให้บริการเสมอ' },
  { value: 'inventory', label: 'ธุรกิจซื้อขายสินค้า', description: 'ใช้รูปแบบต้นทุนสินค้าที่ขายเสมอ' },
  { value: 'both', label: 'แสดงทั้งสองรูปแบบ', description: 'สรุปทั้งต้นทุนการให้บริการและต้นทุนสินค้าที่ขายในหมายเหตุเดียวกัน' }
];
