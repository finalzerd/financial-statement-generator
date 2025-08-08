import React, { useState } from 'react';

interface CompanyInfoFormData {
  companyName: string;
  companyType: 'บริษัทจำกัด' | 'ห้างหุ้นส่วนจำกัด';
  reportingPeriod: string;
  registrationNumber?: string;
  address?: string;
  businessDescription?: string;
}

interface CompanyInfoFormProps {
  onSubmit: (info: CompanyInfoFormData) => void;
  onCancel: () => void;
}

const CompanyInfoForm: React.FC<CompanyInfoFormProps> = ({ onSubmit, onCancel }) => {
  const [formData, setFormData] = useState<CompanyInfoFormData>({
    companyName: '',
    companyType: 'บริษัทจำกัด',
    reportingPeriod: '2024'
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onSubmit(formData);
  };

  return (
    <div className="company-info-form">
      <h3>ข้อมูลบริษัท</h3>
      <form onSubmit={handleSubmit}>
        <div className="form-group">
          <label>ชื่อบริษัท *</label>
          <input
            type="text"
            value={formData.companyName}
            onChange={(e) => setFormData({...formData, companyName: e.target.value})}
            required
            placeholder="เช่น บริษัท เอบีซี จำกัด"
          />
        </div>
        
        <div className="form-group">
          <label>ประเภทบริษัท *</label>
          <select
            value={formData.companyType}
            onChange={(e) => setFormData({...formData, companyType: e.target.value as 'บริษัทจำกัด' | 'ห้างหุ้นส่วนจำกัด'})}
          >
            <option value="บริษัทจำกัด">บริษัทจำกัด</option>
            <option value="ห้างหุ้นส่วนจำกัด">ห้างหุ้นส่วนจำกัด</option>
          </select>
        </div>
        
        <div className="form-group">
          <label>งวดบัญชี *</label>
          <input
            type="text"
            value={formData.reportingPeriod}
            onChange={(e) => setFormData({...formData, reportingPeriod: e.target.value})}
            placeholder="เช่น 2024"
            required
          />
        </div>
        
        <div className="form-group">
          <label>เลขทะเบียนบริษัท</label>
          <input
            type="text"
            value={formData.registrationNumber || ''}
            onChange={(e) => setFormData({...formData, registrationNumber: e.target.value})}
            placeholder="เช่น 0105564000123"
          />
        </div>
        
        <div className="form-group">
          <label>ที่อยู่</label>
          <input
            type="text"
            value={formData.address || ''}
            onChange={(e) => setFormData({...formData, address: e.target.value})}
            placeholder="เช่น 123 ถนนรัชดาภิเษก แขวงลาดยาว เขตจตุจักร กรุงเทพมหานคร 10900"
          />
        </div>
        
        <div className="form-group">
          <label>ลักษณะธุรกิจ</label>
          <textarea
            value={formData.businessDescription || ''}
            onChange={(e) => setFormData({...formData, businessDescription: e.target.value})}
            placeholder="เช่น ขายส่งขายปลีกสินค้าทางการเกษตร เคมีภัณฑ์ ปุ๋ย ยาปราบศัตรูพืช ยาบำรุงพืชและสัตว์ทุกชนิด"
            rows={3}
          />
        </div>
        
        <div className="form-actions">
          <button type="submit">ตกลง</button>
          <button type="button" onClick={onCancel}>ยกเลิก</button>
        </div>
      </form>
    </div>
  );
};

export default CompanyInfoForm;
export type { CompanyInfoFormData };
