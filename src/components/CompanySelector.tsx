import React, { useState, useEffect } from 'react';
import type { Company } from '../types/database';
import { ApiService } from '../services/apiService';
import './CompanySelector.css';

interface CompanySelectorProps {
  onCompanySelected: (company: Company) => void;
  onCompanyCreated: (company: Company) => void;
}

const CompanySelector: React.FC<CompanySelectorProps> = ({
  onCompanySelected,
  onCompanyCreated
}) => {
  const [companies, setCompanies] = useState<Company[]>([]);
  const [loading, setLoading] = useState(true);
  const [showNewCompanyForm, setShowNewCompanyForm] = useState(false);
  const [editingCompany, setEditingCompany] = useState<Company | null>(null);
  const [deletingCompanyId, setDeletingCompanyId] = useState<string | null>(null);
  const [isCreating, setIsCreating] = useState(false);
  const [isUpdating, setIsUpdating] = useState(false);
  const [error, setError] = useState<string | null>(null);
  
  // DBD search states
  const [dbdSearchNumber, setDbdSearchNumber] = useState('');
  const [isSearchingDbd, setIsSearchingDbd] = useState(false);
  const [dbdSearchError, setDbdSearchError] = useState<string | null>(null);
  
  const [newCompany, setNewCompany] = useState({
    name: '',
    type: 'บริษัทจำกัด' as const,
    registrationNumber: '',
    registrationDate: '',
    address: '',
    businessDescription: '',
    taxId: '',
    numberOfShares: 1000000,
    shareValue: 1,
    defaultReportingYear: new Date().getFullYear(),
    periodStartDate: '',
    periodEndDate: '',
    directorName: '',
    approvalMeetingNumber: '',
    approvalMeetingDate: '',
    financialStatementApprovalDate: ''
  });

  const [editForm, setEditForm] = useState({
    name: '',
    type: 'บริษัทจำกัด' as 'ห้างหุ้นส่วนจำกัด' | 'บริษัทจำกัด',
    registrationNumber: '',
    registrationDate: '',
    address: '',
    businessDescription: '',
    taxId: '',
    numberOfShares: 1000000,
    shareValue: 1,
    defaultReportingYear: new Date().getFullYear(),
    periodStartDate: '',
    periodEndDate: '',
    directorName: '',
    approvalMeetingNumber: '',
    approvalMeetingDate: '',
    financialStatementApprovalDate: ''
  });

  // Load companies on component mount
  useEffect(() => {
    loadCompanies();
  }, []);

  const loadCompanies = async () => {
    try {
      setLoading(true);
      setError(null);
      const result = await ApiService.getCompanies();
      setCompanies(result.companies || []);
    } catch (error) {
      console.error('Failed to load companies:', error);
      setError(error instanceof Error ? error.message : 'Failed to load companies');
    } finally {
      setLoading(false);
    }
  };

  const handleCreateCompany = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setIsCreating(true);

    try {
      const result = await ApiService.createCompany({
        ...newCompany,
        defaultReportingYear: newCompany.defaultReportingYear
      });
      
      // Clear form
      setNewCompany({
        name: '',
        type: 'บริษัทจำกัด',
        registrationNumber: '',
        registrationDate: '',
        address: '',
        businessDescription: '',
        taxId: '',
        numberOfShares: 1000000,
        shareValue: 1,
        defaultReportingYear: new Date().getFullYear(),
        periodStartDate: '',
        periodEndDate: '',
        directorName: '',
        approvalMeetingNumber: '',
        approvalMeetingDate: ''
      });
      setShowNewCompanyForm(false);
      
      // Reload companies and notify parent
      await loadCompanies();
      onCompanyCreated(result.company);
      
    } catch (error) {
      setError(error instanceof Error ? error.message : 'Failed to create company');
    } finally {
      setIsCreating(false);
    }
  };

  // DBD search handler
  const handleDbdSearch = async () => {
    if (!dbdSearchNumber || dbdSearchNumber.trim().length !== 13) {
      setDbdSearchError('กรุณาระบุเลขที่จดทะเบียน 13 หลัก');
      return;
    }

    setIsSearchingDbd(true);
    setDbdSearchError(null);

    try {
      const dbdData = await ApiService.searchDbdCompany(dbdSearchNumber.trim());
      
      // Auto-fill form with DBD data
      setNewCompany({
        ...newCompany,
        name: dbdData.name || '',
        type: dbdData.type || 'บริษัทจำกัด',
        registrationNumber: dbdData.registrationNumber || '',
        registrationDate: dbdData.registrationDate || '',
        address: dbdData.address || '',
        businessDescription: dbdData.businessDescription || ''
      });

      // Show success message
      setError(null);
      alert(`✅ พบข้อมูลบริษัท: ${dbdData.name}\n\nกรุณาตรวจสอบข้อมูลและกรอกรายละเอียดเพิ่มเติม (ทุนจดทะเบียน, งวดบัญชี, ฯลฯ)`);
      
    } catch (error) {
      setDbdSearchError(error instanceof Error ? error.message : 'ไม่สามารถค้นหาข้อมูลได้');
    } finally {
      setIsSearchingDbd(false);
    }
  };

  const resetForm = () => {
    setNewCompany({ 
      name: '', 
      type: 'บริษัทจำกัด', 
      registrationNumber: '',
      registrationDate: '',
      address: '', 
      businessDescription: '',
      taxId: '',
      numberOfShares: 1000000,
      shareValue: 1,
      defaultReportingYear: new Date().getFullYear(),
      directorName: '',
      approvalMeetingNumber: '',
      approvalMeetingDate: ''
    });
    setShowNewCompanyForm(false);
    setError(null);
  };

  const handleUpdateCompany = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!editingCompany || !editForm.name.trim()) {
      setError('กรุณาระบุชื่อบริษัท');
      return;
    }

    try {
      setIsUpdating(true);
      setError(null);
      
      const result = await ApiService.updateCompany(editingCompany.id, {
        ...editForm,
        defaultReportingYear: editForm.defaultReportingYear
      });
      
      await loadCompanies();
      onCompanySelected(result.company);
      resetEditForm();
    } catch (error) {
      console.error('Failed to update company:', error);
      setError(error instanceof Error ? error.message : 'ไม่สามารถอัพเดทบริษัทได้');
    } finally {
      setIsUpdating(false);
    }
  };

  const startEditCompany = (company: Company) => {
    setEditingCompany(company);
    setEditForm({
      name: company.name,
      type: company.type,
      registrationNumber: company.registrationNumber || '',
      registrationDate: company.registrationDate || '',
      address: company.address || '',
      businessDescription: company.businessDescription || '',
      taxId: company.taxId || '',
      numberOfShares: company.numberOfShares || 1000000,
      shareValue: company.shareValue || 1,
      defaultReportingYear: company.defaultReportingYear,
      periodStartDate: company.periodStartDate || '',
      periodEndDate: company.periodEndDate || '',
      directorName: company.directorName || '',
      approvalMeetingNumber: company.approvalMeetingNumber || '',
      approvalMeetingDate: company.approvalMeetingDate || '',
      financialStatementApprovalDate: company.financialStatementApprovalDate || ''
    });
    setShowNewCompanyForm(false);
  };

  const resetEditForm = () => {
    setEditingCompany(null);
    setEditForm({
      name: '',
      type: 'บริษัทจำกัด' as 'ห้างหุ้นส่วนจำกัด' | 'บริษัทจำกัด',
      registrationNumber: '',
      registrationDate: '',
      address: '',
      businessDescription: '',
      taxId: '',
      numberOfShares: 1000000,
      shareValue: 1,
      defaultReportingYear: new Date().getFullYear(),
      periodStartDate: '',
      periodEndDate: '',
      directorName: '',
      approvalMeetingNumber: '',
      approvalMeetingDate: '',
      financialStatementApprovalDate: ''
    });
  };

  const handleDeleteCompany = async (company: Company) => {
    const confirmMessage = `คุณแน่ใจหรือไม่ที่จะลบบริษัท "${company.name}"?\n\nการดำเนินการนี้ไม่สามารถย้อนกลับได้`;
    
    if (!window.confirm(confirmMessage)) {
      return;
    }

    try {
      setDeletingCompanyId(company.id);
      setError(null);
      
      await ApiService.deleteCompany(company.id);
      
      // Reload companies list
      await loadCompanies();
      
      // If the deleted company was being edited, close edit form
      if (editingCompany?.id === company.id) {
        resetEditForm();
      }
      
    } catch (error) {
      console.error('Failed to delete company:', error);
      setError(error instanceof Error ? error.message : 'ไม่สามารถลบบริษัทได้');
    } finally {
      setDeletingCompanyId(null);
    }
  };

  return (
    <div className="company-selector">
      <h2>🏢 เลือกบริษัท / Select Company</h2>
      
      {error && (
        <div className="error-message">
          {error}
        </div>
      )}
      
      <div className="section">
        <h3>บริษัทที่มีอยู่ / Existing Companies</h3>
        
        {loading ? (
          <div className="loading">กำลังโหลดข้อมูลบริษัท...</div>
        ) : companies.length > 0 ? (
          <div className="company-list">
            {companies.map(company => (
              <div
                key={company.id}
                className="company-item"
              >
                <div className="company-info" onClick={() => onCompanySelected(company)}>
                  <h4>{company.name}</h4>
                  <p>{company.type}</p>
                  {company.registrationNumber && (
                    <p>เลขที่จดทะเบียน: {company.registrationNumber}</p>
                  )}
                  {company.address && (
                    <p>ที่อยู่: {company.address}</p>
                  )}
                  {company.businessDescription && (
                    <p>ลักษณะธุรกิจ: {company.businessDescription}</p>
                  )}
                </div>
                <div className="company-actions">
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      startEditCompany(company);
                    }}
                    className="btn btn-secondary btn-small"
                  >
                    ✏️ แก้ไข
                  </button>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleDeleteCompany(company);
                    }}
                    disabled={deletingCompanyId === company.id}
                    className="btn btn-danger btn-small"
                  >
                    {deletingCompanyId === company.id ? '🔄 กำลังลบ...' : '🗑️ ลบ'}
                  </button>
                </div>
                <div className="company-meta">
                  <p>ปีรายงาน: {company.defaultReportingYear}</p>
                  <p>สร้างเมื่อ: {new Date(company.createdAt).toLocaleDateString('th-TH')}</p>
                </div>
              </div>
            ))}
          </div>
        ) : (
          <div className="empty-state">
            <h4>ไม่มีบริษัทในระบบ</h4>
            <p>กรุณาสร้างบริษัทใหม่เพื่อเริ่มต้นใช้งาน</p>
          </div>
        )}
        
        <button
          onClick={() => {
            setShowNewCompanyForm(true);
            setEditingCompany(null);
          }}
          className="btn btn-primary"
          style={{ marginTop: '1rem' }}
        >
          ➕ สร้างบริษัทใหม่
        </button>
      </div>

      {/* Edit Company Form */}
      {editingCompany && (
        <div className="section">
          <h3>แก้ไขข้อมูลบริษัท / Edit Company Information</h3>
          
          <form onSubmit={handleUpdateCompany} className="edit-company-form">
            {/* DBD Auto-Fill Section for Edit Form */}
            <div style={{ 
              backgroundColor: '#f0f9ff', 
              padding: '1rem', 
              borderRadius: '8px', 
              marginBottom: '1.5rem',
              border: '1px solid #bfdbfe'
            }}>
              <h4 style={{ marginTop: 0, marginBottom: '0.75rem', color: '#1e40af' }}>
                🔍 อัพเดทข้อมูลจากกรมพัฒนาธุรกิจการค้า
              </h4>
              <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'flex-start' }}>
                <div style={{ flex: 1 }}>
                  <input
                    type="text"
                    value={dbdSearchNumber}
                    onChange={(e) => {
                      setDbdSearchNumber(e.target.value);
                      setDbdSearchError(null);
                    }}
                    placeholder="ระบุเลขที่จดทะเบียน 13 หลัก"
                    disabled={isSearchingDbd || isUpdating}
                    maxLength={13}
                    pattern="[0-9]{13}"
                    style={{ width: '100%', padding: '0.5rem' }}
                  />
                  {dbdSearchError && (
                    <small style={{ color: '#dc2626', display: 'block', marginTop: '0.25rem' }}>
                      {dbdSearchError}
                    </small>
                  )}
                </div>
                <button
                  type="button"
                  onClick={async () => {
                    if (!dbdSearchNumber || dbdSearchNumber.trim().length !== 13) {
                      setDbdSearchError('กรุณาระบุเลขที่จดทะเบียน 13 หลัก');
                      return;
                    }

                    setIsSearchingDbd(true);
                    setDbdSearchError(null);

                    try {
                      const dbdData = await ApiService.searchDbdCompany(dbdSearchNumber.trim());
                      
                      // Auto-fill edit form with DBD data
                      setEditForm({
                        ...editForm,
                        name: dbdData.name || editForm.name,
                        type: dbdData.type || editForm.type,
                        registrationNumber: dbdData.registrationNumber || editForm.registrationNumber,
                        registrationDate: dbdData.registrationDate || editForm.registrationDate,
                        address: dbdData.address || editForm.address,
                        businessDescription: dbdData.businessDescription || editForm.businessDescription
                      });

                      setError(null);
                      alert(`✅ อัพเดทข้อมูลจาก DBD: ${dbdData.name}`);
                      
                    } catch (error) {
                      setDbdSearchError(error instanceof Error ? error.message : 'ไม่สามารถค้นหาข้อมูลได้');
                    } finally {
                      setIsSearchingDbd(false);
                    }
                  }}
                  disabled={isSearchingDbd || isUpdating || !dbdSearchNumber || dbdSearchNumber.length !== 13}
                  style={{ 
                    padding: '0.5rem 1rem',
                    whiteSpace: 'nowrap',
                    backgroundColor: isSearchingDbd || !dbdSearchNumber || dbdSearchNumber.length !== 13 ? '#9ca3af' : '#3b82f6',
                    color: 'white',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: isSearchingDbd || !dbdSearchNumber || dbdSearchNumber.length !== 13 ? 'not-allowed' : 'pointer'
                  }}
                >
                  {isSearchingDbd ? '🔄 กำลังค้นหา...' : '🔍 อัพเดทจาก DBD'}
                </button>
              </div>
              <small style={{ display: 'block', marginTop: '0.5rem', color: '#64748b' }}>
                💡 ใช้เลขที่จดทะเบียนเพื่ออัพเดทข้อมูลล่าสุดจากระบบ DBD
              </small>
            </div>

            <hr style={{ margin: '1.5rem 0', border: 'none', borderTop: '1px solid #e5e7eb' }} />

            <div className="form-group">
              <label htmlFor="edit-company-name">ชื่อบริษัท / Company Name *</label>
              <input
                type="text"
                id="edit-company-name"
                value={editForm.name}
                onChange={e => setEditForm({...editForm, name: e.target.value})}
                disabled={isUpdating}
                required
              />
            </div>
            
            <div className="form-row">
              <div className="form-group">
                <label htmlFor="edit-company-type">ประเภทบริษัท / Company Type *</label>
                <select
                  id="edit-company-type"
                  value={editForm.type}
                  onChange={e => setEditForm({...editForm, type: e.target.value as 'ห้างหุ้นส่วนจำกัด' | 'บริษัทจำกัด'})}
                  disabled={isUpdating}
                  required
                >
                  <option value="บริษัทจำกัด">บริษัทจำกัด</option>
                  <option value="ห้างหุ้นส่วนจำกัด">ห้างหุ้นส่วนจำกัด</option>
                </select>
              </div>
              
              <div className="form-group">
                <label htmlFor="edit-registration-number">เลขที่จดทะเบียน / Registration Number</label>
                <input
                  type="text"
                  id="edit-registration-number"
                  value={editForm.registrationNumber}
                  onChange={e => setEditForm({...editForm, registrationNumber: e.target.value})}
                  disabled={isUpdating}
                />
              </div>
              <div className="form-group">
                <label htmlFor="edit-registration-date">วันที่จดทะเบียน / Registration Date</label>
                <input
                  type="text"
                  id="edit-registration-date"
                  value={editForm.registrationDate}
                  onChange={e => setEditForm({...editForm, registrationDate: e.target.value})}
                  disabled={isUpdating}
                  placeholder="เช่น 17 ก.ค. 2563"
                />
                <small className="form-help">วันที่จดทะเบียนบริษัท (ข้อความอิสระ)</small>
              </div>
              <div className="form-group">
                <label htmlFor="edit-default-year">ปีรายงานเริ่มต้น / Default Reporting Year *</label>
                <input
                  type="number"
                  id="edit-default-year"
                  value={editForm.defaultReportingYear}
                  onChange={e => setEditForm({...editForm, defaultReportingYear: Number(e.target.value) || new Date().getFullYear()})}
                  disabled={isUpdating}
                  min="1900"
                  required
                />
                <small className="form-help">รองรับปีพุทธศักราช (พ.ศ.) ได้ เช่น 2568</small>
              </div>
              <div className="form-group">
                <label htmlFor="edit-period-start-date">วันที่เริ่มต้นงวดบัญชี / Period Start Date</label>
                <input
                  type="text"
                  id="edit-period-start-date"
                  value={editForm.periodStartDate}
                  onChange={e => setEditForm({...editForm, periodStartDate: e.target.value})}
                  disabled={isUpdating}
                  placeholder="เช่น 1 มกราคม"
                />
                <small className="form-help">ค่าเริ่มต้น: 1 มกราคม</small>
              </div>
              <div className="form-group">
                <label htmlFor="edit-period-end-date">วันที่สิ้นสุดงวดบัญชี / Period End Date</label>
                <input
                  type="text"
                  id="edit-period-end-date"
                  value={editForm.periodEndDate}
                  onChange={e => setEditForm({...editForm, periodEndDate: e.target.value})}
                  disabled={isUpdating}
                  placeholder="เช่น 31 ธันวาคม"
                />
                <small className="form-help">ค่าเริ่มต้น: 31 ธันวาคม</small>
              </div>
            </div>
            
            <div className="form-group">
              <label htmlFor="edit-address">ที่อยู่ / Address</label>
              <textarea
                id="edit-address"
                value={editForm.address}
                onChange={e => setEditForm({...editForm, address: e.target.value})}
                disabled={isUpdating}
                placeholder="เช่น 123 ถนนรัชดาภิเษก แขวงลาดยาว เขตจตุจักร กรุงเทพมหานคร 10900"
                rows={2}
              />
              <small className="form-help">ข้อมูลนี้จะปรากฏในหมายเหตุประกอบงบการเงิน หัวข้อ "1.2 สถานที่ตั้ง"</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="edit-business-description">ลักษณะธุรกิจ / Business Description</label>
              <textarea
                id="edit-business-description"
                value={editForm.businessDescription}
                onChange={e => setEditForm({...editForm, businessDescription: e.target.value})}
                disabled={isUpdating}
                placeholder="เช่น ขายส่งขายปลีกสินค้าทางการเกษตร เคมีภัณฑ์ ปุ๋ย ยาปราบศัตรูพืช ยาบำรุงพืชและสัตว์ทุกชนิด"
                rows={3}
              />
              <small className="form-help">ข้อมูลนี้จะปรากฏในหมายเหตุประกอบงบการเงิน หัวข้อ "1.3 ลักษณะธุรกิจและการดำเนินงาน"</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="edit-tax-id">เลขที่ผู้เสียภาษี / Tax ID</label>
              <input
                type="text"
                id="edit-tax-id"
                value={editForm.taxId}
                onChange={e => setEditForm({...editForm, taxId: e.target.value})}
                disabled={isUpdating}
              />
            </div>
            
            {/* Share information - only for บริษัทจำกัด */}
            {editForm.type === 'บริษัทจำกัด' && (
              <>
                <div className="form-group">
                  <label htmlFor="edit-number-of-shares">จำนวนหุ้นสามัญ (หุ้น) *</label>
                  <input
                    type="number"
                    id="edit-number-of-shares"
                    value={editForm.numberOfShares}
                    onChange={e => setEditForm({...editForm, numberOfShares: Number(e.target.value) || 1000000})}
                    disabled={isUpdating}
                    min="1"
                    required
                  />
                  <small className="form-help">จำนวนหุ้นที่จดทะเบียนและออกจำหน่าย</small>
                </div>
                
                <div className="form-group">
                  <label htmlFor="edit-share-value">มูลค่าหุ้นสามัญ (บาทต่อหุ้น) *</label>
                  <input
                    type="number"
                    id="edit-share-value"
                    value={editForm.shareValue}
                    onChange={e => setEditForm({...editForm, shareValue: Number(e.target.value) || 1})}
                    disabled={isUpdating}
                    min="0.01"
                    step="0.01"
                    required
                  />
                  <small className="form-help">มูลค่าต่อหุ้น (Par Value) ตามหนังสือรับรอง</small>
                </div>
              </>
            )}
            
            {/* Director signature fields */}
            <div className="form-group">
              <label htmlFor="edit-director-name">ชื่อกรรมการ</label>
              <input
                type="text"
                id="edit-director-name"
                value={editForm.directorName}
                onChange={e => setEditForm({...editForm, directorName: e.target.value})}
                disabled={isUpdating}
                placeholder="ชื่อกรรมการผู้มีอำนาจลงนาม"
              />
              <small className="form-help">ชื่อกรรมการที่จะแสดงในส่วนท้ายงบการเงิน</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="edit-approval-meeting-number">ครั้งที่ประชุม</label>
              <input
                type="text"
                id="edit-approval-meeting-number"
                value={editForm.approvalMeetingNumber}
                onChange={e => setEditForm({...editForm, approvalMeetingNumber: e.target.value})}
                disabled={isUpdating}
                placeholder="เช่น 1/2567"
              />
              <small className="form-help">ครั้งที่ประชุมสามัญผู้ถือหุ้นที่อนุมัติงบการเงิน</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="edit-approval-meeting-date">วันที่ประชุม</label>
              <input
                type="text"
                id="edit-approval-meeting-date"
                value={editForm.approvalMeetingDate}
                onChange={e => setEditForm({...editForm, approvalMeetingDate: e.target.value})}
                disabled={isUpdating}
                placeholder="เช่น 28 มีนาคม 2567"
              />
              <small className="form-help">วันที่ประชุมอนุมัติงบการเงิน (ข้อความอิสระ)</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="edit-financial-approval-date">วันที่อนุมัติงบการเงิน</label>
              <input
                type="text"
                id="edit-financial-approval-date"
                value={editForm.financialStatementApprovalDate}
                onChange={e => setEditForm({...editForm, financialStatementApprovalDate: e.target.value})}
                disabled={isUpdating}
                placeholder="เช่น 15 มกราคม 2567"
              />
              <small className="form-help">วันที่อนุมัติงบการเงินโดยคณะกรรมการ/ที่ประชุมหุ้นส่วน</small>
            </div>
            
            <div className="form-actions">
              <button
                type="submit"
                disabled={isUpdating || !editForm.name.trim()}
                className="btn btn-success"
              >
                {isUpdating ? '🔄 กำลังอัพเดท...' : '✅ บันทึกการแก้ไข'}
              </button>
              <button
                type="button"
                onClick={resetEditForm}
                disabled={isUpdating}
                className="btn btn-secondary"
              >
                ❌ ยกเลิก
              </button>
            </div>
          </form>
        </div>
      )}

      {/* New Company Form */}
      {showNewCompanyForm && (
        <div className="section">
          <h3>สร้างบริษัทใหม่ / Create New Company</h3>
          
          <form onSubmit={handleCreateCompany} className="new-company-form">
            {/* DBD Auto-Fill Section */}
            <div style={{ 
              backgroundColor: '#f0f9ff', 
              padding: '1rem', 
              borderRadius: '8px', 
              marginBottom: '1.5rem',
              border: '1px solid #bfdbfe'
            }}>
              <h4 style={{ marginTop: 0, marginBottom: '0.75rem', color: '#1e40af' }}>
                🔍 ค้นหาข้อมูลจากกรมพัฒนาธุรกิจการค้า
              </h4>
              <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'flex-start' }}>
                <div style={{ flex: 1 }}>
                  <input
                    type="text"
                    value={dbdSearchNumber}
                    onChange={(e) => {
                      setDbdSearchNumber(e.target.value);
                      setDbdSearchError(null);
                    }}
                    placeholder="ระบุเลขที่จดทะเบียน 13 หลัก (เช่น 0335567001107)"
                    disabled={isSearchingDbd || isCreating}
                    maxLength={13}
                    pattern="[0-9]{13}"
                    style={{ width: '100%', padding: '0.5rem' }}
                  />
                  {dbdSearchError && (
                    <small style={{ color: '#dc2626', display: 'block', marginTop: '0.25rem' }}>
                      {dbdSearchError}
                    </small>
                  )}
                </div>
                <button
                  type="button"
                  onClick={handleDbdSearch}
                  disabled={isSearchingDbd || isCreating || !dbdSearchNumber || dbdSearchNumber.length !== 13}
                  style={{ 
                    padding: '0.5rem 1rem',
                    whiteSpace: 'nowrap',
                    backgroundColor: isSearchingDbd || !dbdSearchNumber || dbdSearchNumber.length !== 13 ? '#9ca3af' : '#3b82f6',
                    color: 'white',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: isSearchingDbd || !dbdSearchNumber || dbdSearchNumber.length !== 13 ? 'not-allowed' : 'pointer'
                  }}
                >
                  {isSearchingDbd ? '🔄 กำลังค้นหา...' : '🔍 ค้นหา'}
                </button>
              </div>
              <small style={{ display: 'block', marginTop: '0.5rem', color: '#64748b' }}>
                💡 กรอกเลขที่จดทะเบียนเพื่อดึงข้อมูลอัตโนมัติจากระบบกรมพัฒนาธุรกิจการค้า (DBD)
              </small>
            </div>

            <hr style={{ margin: '1.5rem 0', border: 'none', borderTop: '1px solid #e5e7eb' }} />
            
            <div className="form-group">
              <label htmlFor="company-name">ชื่อบริษัท / Company Name *</label>
              <input
                type="text"
                id="company-name"
                value={newCompany.name}
                onChange={e => setNewCompany({...newCompany, name: e.target.value})}
                disabled={isCreating}
                required
              />
            </div>
            
            <div className="form-row">
              <div className="form-group">
                <label htmlFor="company-type">ประเภทบริษัท / Company Type *</label>
                <select
                  id="company-type"
                  value={newCompany.type}
                  onChange={e => setNewCompany({...newCompany, type: e.target.value as any})}
                  disabled={isCreating}
                >
                  <option value="บริษัทจำกัด">บริษัทจำกัด</option>
                  <option value="ห้างหุ้นส่วนจำกัด">ห้างหุ้นส่วนจำกัด</option>
                </select>
              </div>
              
              <div className="form-group">
                <label htmlFor="registration-number">เลขที่จดทะเบียน / Registration Number</label>
                <input
                  type="text"
                  id="registration-number"
                  value={newCompany.registrationNumber}
                  onChange={e => setNewCompany({...newCompany, registrationNumber: e.target.value})}
                  disabled={isCreating}
                />
              </div>
              <div className="form-group">
                <label htmlFor="registration-date">วันที่จดทะเบียน / Registration Date</label>
                <input
                  type="text"
                  id="registration-date"
                  value={newCompany.registrationDate}
                  onChange={e => setNewCompany({...newCompany, registrationDate: e.target.value})}
                  disabled={isCreating}
                  placeholder="เช่น 17 ก.ค. 2563"
                />
                <small className="form-help">วันที่จดทะเบียนบริษัท (ข้อความอิสระ)</small>
              </div>
              <div className="form-group">
                <label htmlFor="default-year">ปีรายงานเริ่มต้น / Default Reporting Year *</label>
                <input
                  type="number"
                  id="default-year"
                  value={newCompany.defaultReportingYear}
                  onChange={e => setNewCompany({...newCompany, defaultReportingYear: Number(e.target.value) || new Date().getFullYear()})}
                  disabled={isCreating}
                  min="1900"
                  required
                />
                <small className="form-help">รองรับปีพุทธศักราช (พ.ศ.) ได้ เช่น 2568</small>
              </div>
              <div className="form-group">
                <label htmlFor="period-start-date">วันที่เริ่มต้นงวดบัญชี / Period Start Date</label>
                <input
                  type="text"
                  id="period-start-date"
                  value={newCompany.periodStartDate}
                  onChange={e => setNewCompany({...newCompany, periodStartDate: e.target.value})}
                  disabled={isCreating}
                  placeholder="เช่น 1 มกราคม"
                />
                <small className="form-help">ค่าเริ่มต้น: 1 มกราคม</small>
              </div>
              <div className="form-group">
                <label htmlFor="period-end-date">วันที่สิ้นสุดงวดบัญชี / Period End Date</label>
                <input
                  type="text"
                  id="period-end-date"
                  value={newCompany.periodEndDate}
                  onChange={e => setNewCompany({...newCompany, periodEndDate: e.target.value})}
                  disabled={isCreating}
                  placeholder="เช่น 31 ธันวาคม"
                />
                <small className="form-help">ค่าเริ่มต้น: 31 ธันวาคม</small>
              </div>
            </div>
            
            <div className="form-group">
              <label htmlFor="address">ที่อยู่ / Address</label>
              <input
                type="text"
                id="address"
                value={newCompany.address}
                onChange={e => setNewCompany({...newCompany, address: e.target.value})}
                disabled={isCreating}
                placeholder="เช่น 123 ถนนรัชดาภิเษก แขวงลาดยาว เขตจตุจักร กรุงเทพมหานคร 10900"
              />
              <small className="form-help">ข้อมูลนี้จะปรากฏในหมายเหตุประกอบงบการเงิน หัวข้อ "1.2 สถานที่ตั้ง"</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="business-description">ลักษณะธุรกิจ / Business Description</label>
              <textarea
                id="business-description"
                value={newCompany.businessDescription}
                onChange={e => setNewCompany({...newCompany, businessDescription: e.target.value})}
                disabled={isCreating}
                placeholder="เช่น ขายส่งขายปลีกสินค้าทางการเกษตร เคมีภัณฑ์ ปุ๋ย ยาปราบศัตรูพืช ยาบำรุงพืชและสัตว์ทุกชนิด"
                rows={3}
              />
              <small className="form-help">ข้อมูลนี้จะปรากฏในหมายเหตุประกอบงบการเงิน หัวข้อ "1.3 ลักษณะธุรกิจและการดำเนินงาน"</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="tax-id">เลขที่ผู้เสียภาษี / Tax ID</label>
              <input
                type="text"
                id="tax-id"
                value={newCompany.taxId}
                onChange={e => setNewCompany({...newCompany, taxId: e.target.value})}
                disabled={isCreating}
              />
            </div>
            
            {/* Share information - only for บริษัทจำกัด */}
            {newCompany.type === 'บริษัทจำกัด' && (
              <>
                <div className="form-group">
                  <label htmlFor="number-of-shares">จำนวนหุ้นสามัญ (หุ้น) *</label>
                  <input
                    type="number"
                    id="number-of-shares"
                    value={newCompany.numberOfShares}
                    onChange={e => setNewCompany({...newCompany, numberOfShares: Number(e.target.value) || 1000000})}
                    disabled={isCreating}
                    min="1"
                    required
                  />
                  <small className="form-help">จำนวนหุ้นที่จดทะเบียนและออกจำหน่าย</small>
                </div>
                
                <div className="form-group">
                  <label htmlFor="share-value">มูลค่าหุ้นสามัญ (บาทต่อหุ้น) *</label>
                  <input
                    type="number"
                    id="share-value"
                    value={newCompany.shareValue}
                    onChange={e => setNewCompany({...newCompany, shareValue: Number(e.target.value) || 1})}
                    disabled={isCreating}
                    min="0.01"
                    step="0.01"
                    required
                  />
                  <small className="form-help">มูลค่าต่อหุ้น (Par Value) ตามหนังสือรับรอง</small>
                </div>
              </>
            )}
            
            {/* Director signature fields */}
            <div className="form-group">
              <label htmlFor="director-name">ชื่อกรรมการ</label>
              <input
                type="text"
                id="director-name"
                value={newCompany.directorName}
                onChange={e => setNewCompany({...newCompany, directorName: e.target.value})}
                disabled={isCreating}
                placeholder="ชื่อกรรมการผู้มีอำนาจลงนาม"
              />
              <small className="form-help">ชื่อกรรมการที่จะแสดงในส่วนท้ายงบการเงิน</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="approval-meeting-number">ครั้งที่ประชุม</label>
              <input
                type="text"
                id="approval-meeting-number"
                value={newCompany.approvalMeetingNumber}
                onChange={e => setNewCompany({...newCompany, approvalMeetingNumber: e.target.value})}
                disabled={isCreating}
                placeholder="เช่น 1/2567"
              />
              <small className="form-help">ครั้งที่ประชุมสามัญผู้ถือหุ้นที่อนุมัติงบการเงิน</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="approval-meeting-date">วันที่ประชุม</label>
              <input
                type="text"
                id="approval-meeting-date"
                value={newCompany.approvalMeetingDate}
                onChange={e => setNewCompany({...newCompany, approvalMeetingDate: e.target.value})}
                disabled={isCreating}
                placeholder="เช่น 28 มีนาคม 2567"
              />
              <small className="form-help">วันที่ประชุมอนุมัติงบการเงิน (ข้อความอิสระ)</small>
            </div>
            
            <div className="form-group">
              <label htmlFor="financial-approval-date">วันที่อนุมัติงบการเงิน</label>
              <input
                type="text"
                id="financial-approval-date"
                value={newCompany.financialStatementApprovalDate}
                onChange={e => setNewCompany({...newCompany, financialStatementApprovalDate: e.target.value})}
                disabled={isCreating}
                placeholder="เช่น 15 มกราคม 2567"
              />
              <small className="form-help">วันที่อนุมัติงบการเงินโดยคณะกรรมการ/ที่ประชุมหุ้นส่วน</small>
            </div>
            
            <div className="form-actions">
              <button
                type="submit"
                disabled={isCreating || !newCompany.name.trim()}
                className="btn btn-success"
              >
                {isCreating ? '🔄 กำลังสร้าง...' : '✅ สร้างบริษัท'}
              </button>
              <button
                type="button"
                onClick={resetForm}
                disabled={isCreating}
                className="btn btn-secondary"
              >
                ❌ ยกเลิก
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
};

export default CompanySelector;
