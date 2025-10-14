import { useState, useEffect } from 'react';
import { ApiService } from '../services/apiService';
import type { 
  CompanyAccountMapping, 
  AccountMappingRules, 
  AccountMappingValidation 
} from '../types/accountMapping';
import { 
  STANDARD_NOTE_TYPES, 
  AccountMappingUtils
} from '../types/accountMapping';
import type { DetailOneMode } from '../types/detailSettings';
import { DETAIL_ONE_MODE_OPTIONS } from '../types/detailSettings';
import './AccountMappingManager.css';

interface AccountMappingManagerProps {
  companyId: string;
  onMappingsChanged?: () => void;
  trialBalanceData?: any[]; // For real-time validation
}

interface MappingFormData {
  noteType: string;
  noteNumber: number;
  noteTitle: string;
  ranges: Array<{ from: string; to: string }>;
  includes: string;
  excludes: string;
  isActive: boolean;
  // Optional sub-category rule editing (currently only for cash)
  subCategoryRules?: {
    cash?: {
      ranges: Array<{ from: string; to: string }>;
      includes: string;
      excludes: string;
    };
    bankDeposits?: {
      ranges: Array<{ from: string; to: string }>;
      includes: string;
      excludes: string;
    };
  } | null;
}

export function AccountMappingManager({ 
  companyId, 
  onMappingsChanged, 
  trialBalanceData = [] 
}: AccountMappingManagerProps) {
  const [mappings, setMappings] = useState<CompanyAccountMapping[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [editingMapping, setEditingMapping] = useState<string | null>(null);
  const [formData, setFormData] = useState<MappingFormData | null>(null);
  const [validation, setValidation] = useState<AccountMappingValidation | null>(null);
  const [addNoteType, setAddNoteType] = useState<string>('');
  const [detailOneMode, setDetailOneMode] = useState<DetailOneMode>('auto');
  const [detailSettingsLoaded, setDetailSettingsLoaded] = useState(false);
  const [detailSettingsError, setDetailSettingsError] = useState<string | null>(null);
  const [detailModeSaving, setDetailModeSaving] = useState(false);

  // Load mappings on component mount
  useEffect(() => {
    loadMappings();
    loadDetailSettings();
  }, [companyId]);

  // Validate mappings when trial balance data changes
  useEffect(() => {
    if (trialBalanceData.length > 0 && mappings.length > 0) {
      validateMappings();
    }
  }, [trialBalanceData, mappings]);

  const loadMappings = async () => {
    try {
      setLoading(true);
      const data = await ApiService.getCompanyAccountMappings(companyId);
      setMappings(data);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load account mappings');
    } finally {
      setLoading(false);
    }
  };

  const loadDetailSettings = async () => {
    try {
      setDetailSettingsLoaded(false);
      const settings = await ApiService.getCompanyDetailSettings(companyId);
      setDetailOneMode(settings.detailOneMode);
      setDetailSettingsError(null);
    } catch (err) {
      console.error('Failed to load detail settings:', err);
      setDetailSettingsError(err instanceof Error ? err.message : 'Failed to load detail settings');
      setDetailOneMode('auto');
    } finally {
      setDetailSettingsLoaded(true);
    }
  };

  const validateMappings = async () => {
    try {
      const validationResult = await ApiService.validateAccountMappings(
        companyId, 
        trialBalanceData
      );
      setValidation(validationResult);
    } catch (err) {
      console.error('Validation failed:', err);
    }
  };

  const startEditing = (mapping: CompanyAccountMapping) => {
    setEditingMapping(mapping.noteType);
    setFormData({
      noteType: mapping.noteType,
      noteNumber: mapping.noteNumber || STANDARD_NOTE_TYPES[mapping.noteType as keyof typeof STANDARD_NOTE_TYPES]?.noteNumber || 0,
      noteTitle: mapping.noteTitle || STANDARD_NOTE_TYPES[mapping.noteType as keyof typeof STANDARD_NOTE_TYPES]?.noteTitle || '',
      ranges: mapping.accountRanges.ranges?.map(r => ({ 
        from: r.from.toString(), 
        to: r.to.toString() 
      })) || [{ from: '', to: '' }],
      includes: mapping.accountRanges.includes?.join(', ') || '',
      excludes: mapping.accountRanges.excludes?.join(', ') || '',
      isActive: mapping.isActive,
      subCategoryRules: mapping.subCategoryRules?.cash ? {
        cash: {
          ranges: mapping.subCategoryRules.cash.cash?.ranges?.map((r: any) => ({ from: r.from.toString(), to: r.to.toString() })) || [{ from: '', to: '' }],
          includes: mapping.subCategoryRules.cash.cash?.includes?.join(', ') || '',
          excludes: mapping.subCategoryRules.cash.cash?.excludes?.join(', ') || ''
        },
        bankDeposits: {
          ranges: mapping.subCategoryRules.cash.bankDeposits?.ranges?.map((r: any) => ({ from: r.from.toString(), to: r.to.toString() })) || [{ from: '', to: '' }],
          includes: mapping.subCategoryRules.cash.bankDeposits?.includes?.join(', ') || '',
          excludes: mapping.subCategoryRules.cash.bankDeposits?.excludes?.join(', ') || ''
        }
      } : null
    });
  };

  const cancelEditing = () => {
    setEditingMapping(null);
    setFormData(null);
  };

  const saveMapping = async () => {
    if (!formData) return;

    try {
      // Parse form data into proper format
      const accountRanges: AccountMappingRules = {
        ranges: formData.ranges
          .filter(r => r.from && r.to)
          .map(r => ({ from: parseFloat(r.from), to: parseFloat(r.to) }))
          .filter(r => !isNaN(r.from) && !isNaN(r.to)),
        includes: formData.includes
          .split(',')
          .map(s => s.trim())
          .filter(s => s)
          .map(s => parseFloat(s))
          .filter(n => !isNaN(n)),
        excludes: formData.excludes
          .split(',')
          .map(s => s.trim())
          .filter(s => s)
          .map(s => parseFloat(s))
          .filter(n => !isNaN(n))
      };

      // No sub-categories in simplified version

      // Validate rules
      const ruleValidation = AccountMappingUtils.validateMappingRules(accountRanges);
      if (!ruleValidation.isValid) {
        setError(ruleValidation.errors.join(', '));
        return;
      }

      // Build subCategoryRules structure (only for cash)
      let subCategoryRules: any = null;
      if (formData.noteType === 'cash' && formData.subCategoryRules) {
        const buildSub = (sc: { ranges: { from: string; to: string }[]; includes: string; excludes: string }) => {
          const r = sc.ranges
            .filter(r => r.from && r.to)
            .map(r => ({ from: parseFloat(r.from), to: parseFloat(r.to) }))
            .filter(r => !isNaN(r.from) && !isNaN(r.to));
          const inc = sc.includes
            .split(',')
            .map(s => s.trim())
            .filter(s => s)
            .map(s => parseFloat(s))
            .filter(n => !isNaN(n));
          const exc = sc.excludes
            .split(',')
            .map(s => s.trim())
            .filter(s => s)
            .map(s => parseFloat(s))
            .filter(n => !isNaN(n));
          return { ...(r.length ? { ranges: r } : {}), ...(inc.length ? { includes: inc } : {}), ...(exc.length ? { excludes: exc } : {}) };
        };
        const cashRules = buildSub(formData.subCategoryRules.cash!);
        const bankRules = buildSub(formData.subCategoryRules.bankDeposits!);
        if (Object.keys(cashRules).length || Object.keys(bankRules).length) {
          subCategoryRules = { cash: { cash: cashRules, bankDeposits: bankRules } };
        }
      }

      await ApiService.updateAccountMapping(companyId, formData.noteType, {
        noteNumber: formData.noteNumber,
        noteTitle: formData.noteTitle,
        accountRanges,
        isActive: formData.isActive,
        subCategoryRules
      });

      await loadMappings();
      setEditingMapping(null);
      setFormData(null);
      setError(null);
      onMappingsChanged?.();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save mapping');
    }
  };

  const resetToDefaults = async () => {
    if (!confirm('Reset all account mappings to default values? This will overwrite all customizations.')) {
      return;
    }

    try {
      await ApiService.resetAccountMappingsToDefault(companyId);
      await loadMappings();
      await loadDetailSettings();
      onMappingsChanged?.();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to reset mappings');
    }
  };

  // Create a missing mapping (e.g., asset_short_term_loans) with sensible defaults
  const handleAddMapping = async () => {
    if (!addNoteType) return;
    try {
      const def = (STANDARD_NOTE_TYPES as any)[addNoteType];
      if (!def) {
        setError('Unknown note type');
        return;
      }

      // Default accountRanges for new asset-side notes; others start empty
      let accountRanges: AccountMappingRules = { ranges: [] } as any;
      if (addNoteType === 'asset_short_term_loans') {
        accountRanges = { includes: [1141] } as any; // เงินให้กู้ยืมระยะสั้น (asset)
      } else if (addNoteType === 'asset_long_term_loans') {
        accountRanges = { includes: [1710] } as any; // เงินให้กู้ยืมระยะยาว (asset)
      } else if (addNoteType === 'hire_purchase_creditors') {
        accountRanges = { includes: [2015, 1644.2, 1644.1] } as any; // เจ้าหนี้ตามสัญญาเช่าซื้อ
      } else if (addNoteType === 'other_income') {
        accountRanges = { ranges: [{ from: 4110, to: 4999 }] } as any;
      }

      await ApiService.createAccountMapping(companyId, {
        noteType: addNoteType,
        noteNumber: def.noteNumber,
        noteTitle: def.noteTitle,
        accountRanges,
        isActive: true
      });

      setAddNoteType('');
      await loadMappings();
      onMappingsChanged?.();
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to add mapping');
    }
  };

  const addRange = () => {
    if (formData) {
      setFormData({
        ...formData,
        ranges: [...formData.ranges, { from: '', to: '' }]
      });
    }
  };

  const removeRange = (index: number) => {
    if (formData) {
      setFormData({
        ...formData,
        ranges: formData.ranges.filter((_, i) => i !== index)
      });
    }
  };

  const updateRange = (index: number, field: 'from' | 'to', value: string) => {
    if (formData) {
      const newRanges = [...formData.ranges];
      newRanges[index] = { ...newRanges[index], [field]: value };
      setFormData({ ...formData, ranges: newRanges });
    }
  };

  const addSubRange = (category: 'cash' | 'bankDeposits') => {
    if (!formData) return;
    const sc = formData.subCategoryRules || {
      cash: { ranges: [{ from: '', to: '' }], includes: '', excludes: '' },
      bankDeposits: { ranges: [{ from: '', to: '' }], includes: '', excludes: '' }
    };
    sc[category]!.ranges = [...sc[category]!.ranges, { from: '', to: '' }];
    setFormData({ ...formData, subCategoryRules: sc });
  };

  const updateSubRange = (category: 'cash' | 'bankDeposits', index: number, field: 'from' | 'to', value: string) => {
    if (!formData) return;
    if (!formData.subCategoryRules) return;
    const sc = { ...formData.subCategoryRules } as any;
    const ranges = [...sc[category]!.ranges];
    ranges[index] = { ...ranges[index], [field]: value };
    sc[category] = { ...sc[category], ranges };
    setFormData({ ...formData, subCategoryRules: sc });
  };

  const removeSubRange = (category: 'cash' | 'bankDeposits', index: number) => {
    if (!formData || !formData.subCategoryRules) return;
    const sc = { ...formData.subCategoryRules } as any;
    sc[category]!.ranges = sc[category]!.ranges.filter((_: any, i: number) => i !== index);
    if (sc[category]!.ranges.length === 0) sc[category]!.ranges = [{ from: '', to: '' }];
    setFormData({ ...formData, subCategoryRules: sc });
  };

  const handleDetailModeChange = async (mode: DetailOneMode) => {
    if (mode === detailOneMode && detailSettingsLoaded) {
      return;
    }

    const previousMode = detailOneMode;
    setDetailOneMode(mode);
    setDetailModeSaving(true);
    setDetailSettingsError(null);

    try {
      const updated = await ApiService.updateCompanyDetailSettings(companyId, { detailOneMode: mode });
      setDetailOneMode(updated.detailOneMode);
      if (updated.persisted === false) {
        setDetailSettingsError('ไม่สามารถบันทึกการตั้งค่ากับเซิร์ฟเวอร์ได้ (404). กรุณารีสตาร์ทหรืออัปเดตเซิร์ฟเวอร์ แล้วลองอีกครั้ง');
      } else {
        setDetailSettingsError(null);
      }
    } catch (err) {
      console.error('Failed to update detail settings:', err);
      setDetailOneMode(previousMode);
      setDetailSettingsError(err instanceof Error ? err.message : 'Failed to update detail settings');
    } finally {
      setDetailModeSaving(false);
    }
  };

  // No sub-category helpers in simplified version

  if (loading) return <div className="loading">Loading account mappings...</div>;

  return (
    <div className="account-mapping-manager">
      <div className="header">
        <h2>Account Code Mappings</h2>
        <div className="header-actions">
          <button onClick={resetToDefaults} className="btn-secondary">
            Reset to Defaults
          </button>
          <button onClick={validateMappings} className="btn-primary">
            Validate Mappings
          </button>
          {/* Add Missing Mapping control */}
          {(() => {
            const existing = new Set(mappings.map(m => m.noteType));
            const allTypes = Object.keys(STANDARD_NOTE_TYPES);
            const missing = allTypes.filter(t => !existing.has(t));
            if (missing.length === 0) return null;
            return (
              <div className="add-mapping-inline">
                <select value={addNoteType} onChange={e => setAddNoteType(e.target.value)}>
                  <option value="">Add mapping…</option>
                  {missing.map(t => (
                    <option key={t} value={t}>
                      {`Note ${STANDARD_NOTE_TYPES[t as keyof typeof STANDARD_NOTE_TYPES]?.noteNumber || ''}: ${STANDARD_NOTE_TYPES[t as keyof typeof STANDARD_NOTE_TYPES]?.noteTitle || t}`}
                    </option>
                  ))}
                </select>
                <button className="btn-primary" onClick={handleAddMapping} disabled={!addNoteType}>
                  Add
                </button>
              </div>
            );
          })()}
        </div>
      </div>

      <div className="detail-settings-card">
        <div className="detail-settings-header">
          <h3>การตั้งค่าหมายเหตุ รายละเอียดประกอบที่ 1 (DT1)</h3>
          {detailModeSaving && (
            <span className="detail-settings-status">กำลังบันทึก…</span>
          )}
        </div>
        <p className="detail-settings-description">
          กำหนดรูปแบบที่ต้องการสำหรับ &quot;รายละเอียดประกอบที่ 1&quot; โดยเลือกว่าจะใช้โครงสร้างสำหรับธุรกิจบริการ ธุรกิจสินค้าคงเหลือ หรือแสดงทั้งสองรูปแบบ
        </p>
        {!detailSettingsLoaded ? (
          <div className="detail-settings-loading">กำลังโหลดการตั้งค่า…</div>
        ) : (
          <>
            {detailSettingsError && (
              <div className="detail-settings-error">{detailSettingsError}</div>
            )}
            <div className="detail-mode-options">
              {DETAIL_ONE_MODE_OPTIONS.map(option => (
                <label
                  key={option.value}
                  className={`detail-mode-option ${detailOneMode === option.value ? 'active' : ''}`}
                >
                  <input
                    type="radio"
                    name="detail-one-mode"
                    value={option.value}
                    checked={detailOneMode === option.value}
                    onChange={() => handleDetailModeChange(option.value)}
                    disabled={detailModeSaving}
                  />
                  <div className="detail-mode-content">
                    <span className="detail-mode-label">{option.label}</span>
                    <span className="detail-mode-description">{option.description}</span>
                  </div>
                </label>
              ))}
            </div>
          </>
        )}
      </div>

      {error && (
        <div className="error-message">
          {error}
        </div>
      )}

      {validation && (
        <div className={`validation-summary ${validation.isValid ? 'valid' : 'invalid'}`}>
          <h3>Mapping Validation Results</h3>
          <div className="validation-stats">
            <span>Coverage: {validation.mappingCoverage.coveragePercentage.toFixed(1)}%</span>
            <span>Mapped: {validation.mappingCoverage.mappedAccounts}</span>
            <span>Unmapped: {validation.mappingCoverage.unmappedAccounts}</span>
          </div>
          
          {validation.errors.length > 0 && (
            <div className="validation-errors">
              <h4>Errors:</h4>
              <ul>
                {validation.errors.map((error, i) => (
                  <li key={i}>{error}</li>
                ))}
              </ul>
            </div>
          )}

          {validation.warnings.length > 0 && (
            <div className="validation-warnings">
              <h4>Warnings:</h4>
              <ul>
                {validation.warnings.map((warning, i) => (
                  <li key={i}>{warning}</li>
                ))}
              </ul>
            </div>
          )}

          {validation.conflictingAccounts.length > 0 && (
            <div className="conflicting-accounts">
              <h4>Conflicting Accounts (mapped to multiple notes):</h4>
              <ul>
                {validation.conflictingAccounts.slice(0, 10).map((account, i) => (
                  <li key={i}>
                    <div className="conflict-account-line">
                      <span className="account-code">{account.accountCode}</span>
                      <span className="account-name">{account.accountName}</span>
                    </div>
                    <div className="conflict-note-tags">
                      {account.noteTypes.map((note, idx) => (
                        <span key={idx} className="note-tag">{note}</span>
                      ))}
                    </div>
                  </li>
                ))}
                {validation.conflictingAccounts.length > 10 && (
                  <li>... and {validation.conflictingAccounts.length - 10} more</li>
                )}
              </ul>
            </div>
          )}

          {validation.unmappedAccounts.length > 0 && (
            <div className="unmapped-accounts">
              <h4>Unmapped Accounts:</h4>
              <ul>
                {validation.unmappedAccounts.slice(0, 10).map((account, i) => (
                  <li key={i}>
                    {account.accountCode}: {account.accountName} 
                    (Balance: {account.balance.toLocaleString()})
                  </li>
                ))}
                {validation.unmappedAccounts.length > 10 && (
                  <li>... and {validation.unmappedAccounts.length - 10} more</li>
                )}
              </ul>
            </div>
          )}
        </div>
      )}

      <div className="mappings-grid">
        {mappings.map(mapping => (
          <div key={mapping.noteType} className="mapping-card">
            <div className="mapping-header">
              <h3>Note {mapping.noteNumber}: {mapping.noteTitle}</h3>
              <div className="mapping-actions">
                {editingMapping === mapping.noteType ? (
                  <>
                    <button onClick={saveMapping} className="btn-success">Save</button>
                    <button onClick={cancelEditing} className="btn-secondary">Cancel</button>
                  </>
                ) : (
                  <button onClick={() => startEditing(mapping)} className="btn-primary">
                    Edit
                  </button>
                )}
              </div>
            </div>

            {editingMapping === mapping.noteType && formData ? (
              <div className="mapping-form">
                <div className="form-group">
                  <label>Note Number:</label>
                  <input
                    type="number"
                    value={formData.noteNumber}
                    onChange={(e) => setFormData({ ...formData, noteNumber: parseInt(e.target.value) || 0 })}
                  />
                </div>

                <div className="form-group">
                  <label>Note Title:</label>
                  <input
                    type="text"
                    value={formData.noteTitle}
                    onChange={(e) => setFormData({ ...formData, noteTitle: e.target.value })}
                  />
                </div>

                <div className="form-group">
                  <label>Account Ranges:</label>
                  {formData.ranges.map((range, index) => (
                    <div key={index} className="range-input">
                      <input
                        type="number"
                        placeholder="From"
                        value={range.from}
                        onChange={(e) => updateRange(index, 'from', e.target.value)}
                      />
                      <span>to</span>
                      <input
                        type="number"
                        placeholder="To"
                        value={range.to}
                        onChange={(e) => updateRange(index, 'to', e.target.value)}
                      />
                      <button 
                        onClick={() => removeRange(index)} 
                        className="btn-danger-small"
                        disabled={formData.ranges.length === 1}
                      >
                        Remove
                      </button>
                    </div>
                  ))}
                  <button onClick={addRange} className="btn-secondary-small">
                    Add Range
                  </button>
                </div>

                <div className="form-group">
                  <label>Include Specific Accounts (comma-separated):</label>
                  <input
                    type="text"
                    placeholder="1150, 1180, 1200"
                    value={formData.includes}
                    onChange={(e) => setFormData({ ...formData, includes: e.target.value })}
                  />
                </div>

                <div className="form-group">
                  <label>Exclude Specific Accounts (comma-separated):</label>
                  <input
                    type="text"
                    placeholder="1025, 1095"
                    value={formData.excludes}
                    onChange={(e) => setFormData({ ...formData, excludes: e.target.value })}
                  />
                </div>

                {formData.noteType === 'cash' && (
                  <div className="form-group">
                    <label>Cash Sub-Categories</label>
                    {(() => {
                      if (!formData.subCategoryRules) {
                        return (
                          <button
                            type="button"
                            className="btn-secondary-small"
                            onClick={() => setFormData({
                              ...formData,
                              subCategoryRules: {
                                cash: { ranges: [{ from: '', to: '' }], includes: '', excludes: '' },
                                bankDeposits: { ranges: [{ from: '', to: '' }], includes: '', excludes: '' }
                              }
                            })}
                          >Enable Sub-Categories</button>
                        );
                      }
                      const sc = formData.subCategoryRules;
                      return (
                        <div className="sub-category-panels">
                          <div className="sub-cat-panel">
                            <h4>เงินสด (Cash)</h4>
                            {sc.cash?.ranges.map((r, idx) => (
                              <div key={idx} className="range-input">
                                <input type="number" placeholder="From" value={r.from} onChange={(e) => updateSubRange('cash', idx, 'from', e.target.value)} />
                                <span>to</span>
                                <input type="number" placeholder="To" value={r.to} onChange={(e) => updateSubRange('cash', idx, 'to', e.target.value)} />
                                <button className="btn-danger-small" onClick={() => removeSubRange('cash', idx)} disabled={sc.cash!.ranges.length === 1}>Remove</button>
                              </div>
                            ))}
                            <button className="btn-secondary-small" onClick={() => addSubRange('cash')}>Add Cash Range</button>
                            <div className="sub-inline-inputs">
                              <input type="text" placeholder="Cash includes" value={sc.cash?.includes || ''} onChange={(e) => setFormData({ ...formData, subCategoryRules: { ...sc, cash: { ...sc.cash!, includes: e.target.value } } })} />
                              <input type="text" placeholder="Cash excludes" value={sc.cash?.excludes || ''} onChange={(e) => setFormData({ ...formData, subCategoryRules: { ...sc, cash: { ...sc.cash!, excludes: e.target.value } } })} />
                            </div>
                          </div>
                          <div className="sub-cat-panel">
                            <h4>เงินฝากธนาคาร (Bank Deposits)</h4>
                            {sc.bankDeposits?.ranges.map((r, idx) => (
                              <div key={idx} className="range-input">
                                <input type="number" placeholder="From" value={r.from} onChange={(e) => updateSubRange('bankDeposits', idx, 'from', e.target.value)} />
                                <span>to</span>
                                <input type="number" placeholder="To" value={r.to} onChange={(e) => updateSubRange('bankDeposits', idx, 'to', e.target.value)} />
                                <button className="btn-danger-small" onClick={() => removeSubRange('bankDeposits', idx)} disabled={sc.bankDeposits!.ranges.length === 1}>Remove</button>
                              </div>
                            ))}
                            <button className="btn-secondary-small" onClick={() => addSubRange('bankDeposits')}>Add Bank Range</button>
                            <div className="sub-inline-inputs">
                              <input type="text" placeholder="Bank includes" value={sc.bankDeposits?.includes || ''} onChange={(e) => setFormData({ ...formData, subCategoryRules: { ...sc, bankDeposits: { ...sc.bankDeposits!, includes: e.target.value } } })} />
                              <input type="text" placeholder="Bank excludes" value={sc.bankDeposits?.excludes || ''} onChange={(e) => setFormData({ ...formData, subCategoryRules: { ...sc, bankDeposits: { ...sc.bankDeposits!, excludes: e.target.value } } })} />
                            </div>
                          </div>
                        </div>
                      );
                    })()}
                  </div>
                )}

                <div className="form-group">
                  <label>
                    <input
                      type="checkbox"
                      checked={formData.isActive}
                      onChange={(e) => setFormData({ ...formData, isActive: e.target.checked })}
                    />
                    Active
                  </label>
                </div>
              </div>
            ) : (
              <div className="mapping-display">
                <div className="mapping-rules">
                  <strong>Mapping Rules:</strong>
                  <div>{AccountMappingUtils.describeMappingRules(mapping.accountRanges)}</div>
                </div>
                <div className="mapping-status">
                  <span className={`status ${mapping.isActive ? 'active' : 'inactive'}`}>
                    {mapping.isActive ? 'Active' : 'Inactive'}
                  </span>
                </div>
                {trialBalanceData.length > 0 && (
                  <div className="matching-accounts">
                    <strong>Matching Accounts:</strong>
                    <span>
                      {AccountMappingUtils.getMatchingAccounts(trialBalanceData, mapping.accountRanges).length} accounts
                    </span>
                  </div>
                )}
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}
