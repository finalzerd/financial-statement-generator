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

  // Load mappings on component mount
  useEffect(() => {
    loadMappings();
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
      isActive: mapping.isActive
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
          .map(r => ({ from: parseInt(r.from), to: parseInt(r.to) })),
        includes: formData.includes
          .split(',')
          .map(s => s.trim())
          .filter(s => s)
          .map(s => parseInt(s))
          .filter(n => !isNaN(n)),
        excludes: formData.excludes
          .split(',')
          .map(s => s.trim())
          .filter(s => s)
          .map(s => parseInt(s))
          .filter(n => !isNaN(n))
      };

      // Validate rules
      const ruleValidation = AccountMappingUtils.validateMappingRules(accountRanges);
      if (!ruleValidation.isValid) {
        setError(ruleValidation.errors.join(', '));
        return;
      }

      await ApiService.updateAccountMapping(companyId, formData.noteType, {
        noteNumber: formData.noteNumber,
        noteTitle: formData.noteTitle,
        accountRanges,
        isActive: formData.isActive
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
      onMappingsChanged?.();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to reset mappings');
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
        </div>
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
