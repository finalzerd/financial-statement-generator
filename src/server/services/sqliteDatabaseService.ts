import { SQLiteConfig } from '../config/sqlite-database';
import type { Company, TrialBalanceSet, TrialBalanceEntry, GeneratedStatement } from '../../types/database';
import { createHash } from 'crypto';

export class SQLiteDatabaseService {
  
  // Company operations
  async createCompany(company: Omit<Company, 'id' | 'createdAt' | 'updatedAt'>): Promise<Company> {
    const db = SQLiteConfig.getDatabase();
    
    const result = await (db as any).runAsync(
      `INSERT INTO companies (name, thai_name, registration_number, company_type, address, business_type, phone, email)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [company.name, company.type, company.registrationNumber, company.type, 
       company.address, company.businessDescription, '', '']
    );

    const insertedCompany = await (db as any).getAsync(
      'SELECT * FROM companies WHERE id = ?',
      [result.lastID]
    );

    return {
      ...insertedCompany,
      id: insertedCompany.id.toString(),
      businessDescription: insertedCompany.business_type,
      createdAt: new Date(insertedCompany.created_at),
      updatedAt: new Date(insertedCompany.updated_at)
    } as Company;
  }

  async getCompany(id: string): Promise<Company | null> {
    const db = SQLiteConfig.getDatabase();
    
    const company = await (db as any).getAsync(
      'SELECT * FROM companies WHERE id = ?',
      [parseInt(id)]
    );

    if (!company) return null;

    return {
      ...company,
      id: company.id.toString(),
      businessDescription: company.business_type,
      createdAt: new Date(company.created_at),
      updatedAt: new Date(company.updated_at)
    } as Company;
  }

  async getAllCompanies(): Promise<Company[]> {
    const db = SQLiteConfig.getDatabase();
    
    const companies = await (db as any).allAsync('SELECT * FROM companies ORDER BY created_at DESC');
    return companies.map((company: any) => ({
      ...company,
      id: company.id.toString(),
      businessDescription: company.business_type,
      createdAt: new Date(company.created_at),
      updatedAt: new Date(company.updated_at)
    })) as Company[];
  }

  async updateCompany(id: string, updates: Partial<Company>): Promise<Company | null> {
    const db = SQLiteConfig.getDatabase();
    
    // Build dynamic update query
    const setClause = Object.keys(updates).map(key => `${key} = ?`).join(', ');
    const values = Object.values(updates);
    
    await (db as any).runAsync(
      `UPDATE companies SET ${setClause}, updated_at = CURRENT_TIMESTAMP WHERE id = ?`,
      [...values, parseInt(id)]
    );

    return this.getCompany(id);
  }

  // Trial Balance operations
  async saveTrialBalanceSet(
    companyId: string, 
    fileName: string, 
    entries: Array<Omit<TrialBalanceEntry, 'id' | 'trialBalanceSetId'>>,
    metadata?: { 
      reportingYear?: number; 
      reportingPeriod?: string; 
      processingType?: 'single-year' | 'multi-year';
    }
  ): Promise<{ trialBalanceSet: TrialBalanceSet; entries: TrialBalanceEntry[] }> {
    const db = SQLiteConfig.getDatabase();
    
    // Generate file hash for duplicate detection
    const fileHash = this.generateFileHash(fileName, entries);

    // Check for duplicates
    const existing = await (db as any).getAsync(
      'SELECT id FROM trial_balance_sets WHERE file_hash = ?',
      [fileHash]
    );

    if (existing) {
      throw new Error('This trial balance file has already been uploaded');
    }

    await (db as any).runAsync('BEGIN TRANSACTION');
    
    try {
      // Insert trial balance set
      const setResult = await (db as any).runAsync(
        `INSERT INTO trial_balance_sets (company_id, file_name, file_hash, period_start, period_end, currency, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [parseInt(companyId), fileName, fileHash, 
         metadata?.reportingYear?.toString(), metadata?.reportingPeriod, 'THB', '']
      );

      const trialBalanceSetId = setResult.lastID;

      // Insert entries
      const insertedEntries: TrialBalanceEntry[] = [];
      for (const entry of entries) {
        const entryResult = await (db as any).runAsync(
          `INSERT INTO trial_balance_entries 
           (trial_balance_set_id, account_code, account_name, debit_amount, credit_amount, balance, account_type)
           VALUES (?, ?, ?, ?, ?, ?, ?)`,
          [trialBalanceSetId, entry.accountCode, entry.accountName, 
           entry.debitAmount, entry.creditAmount, entry.calculatedBalance, entry.accountCategory]
        );

        const insertedEntry = await (db as any).getAsync(
          'SELECT * FROM trial_balance_entries WHERE id = ?',
          [entryResult.lastID]
        );
        
        insertedEntries.push({
          id: insertedEntry.id.toString(),
          trialBalanceSetId: insertedEntry.trial_balance_set_id.toString(),
          accountCode: insertedEntry.account_code,
          accountName: insertedEntry.account_name,
          previousBalance: insertedEntry.previous_balance || 0,
          currentBalance: insertedEntry.current_balance || 0,
          debitAmount: insertedEntry.debit_amount,
          creditAmount: insertedEntry.credit_amount,
          calculatedBalance: insertedEntry.balance,
          accountCategory: insertedEntry.account_type
        });
      }

      await (db as any).runAsync('COMMIT');

      const trialBalanceSet = await (db as any).getAsync(
        'SELECT * FROM trial_balance_sets WHERE id = ?',
        [trialBalanceSetId]
      );

      return {
        trialBalanceSet: {
          id: trialBalanceSet.id.toString(),
          companyId: trialBalanceSet.company_id.toString(),
          reportingYear: parseInt(trialBalanceSet.period_start) || new Date().getFullYear(),
          reportingPeriod: trialBalanceSet.period_end || 'Annual',
          processingType: metadata?.processingType || 'single-year',
          fileName: trialBalanceSet.file_name,
          fileHash: trialBalanceSet.file_hash,
          uploadedAt: new Date(trialBalanceSet.upload_date),
          status: 'processed',
          metadata: {
            totalEntries: insertedEntries.length,
            hasInventory: insertedEntries.some(e => e.accountCode === '1510'),
            hasPPE: insertedEntries.some(e => e.accountCode.startsWith('15')),
            accountCodeRange: {
              min: Math.min(...insertedEntries.map(e => parseInt(e.accountCode))).toString(),
              max: Math.max(...insertedEntries.map(e => parseInt(e.accountCode))).toString()
            }
          }
        } as TrialBalanceSet,
        entries: insertedEntries
      };
    } catch (error) {
      await (db as any).runAsync('ROLLBACK');
      throw error;
    }
  }

  async getTrialBalanceSets(companyId: string): Promise<TrialBalanceSet[]> {
    const db = SQLiteConfig.getDatabase();
    
    const sets = await (db as any).allAsync(
      'SELECT * FROM trial_balance_sets WHERE company_id = ? ORDER BY upload_date DESC',
      [parseInt(companyId)]
    );

    return sets.map((set: any) => ({
      id: set.id.toString(),
      companyId: set.company_id.toString(),
      reportingYear: parseInt(set.period_start) || new Date().getFullYear(),
      reportingPeriod: set.period_end || 'Annual',
      processingType: 'single-year',
      fileName: set.file_name,
      fileHash: set.file_hash,
      uploadedAt: new Date(set.upload_date),
      status: 'processed',
      metadata: {
        totalEntries: 0,
        hasInventory: false,
        hasPPE: false,
        accountCodeRange: { min: '0', max: '9999' }
      }
    })) as TrialBalanceSet[];
  }

  async getTrialBalanceEntries(trialBalanceSetId: string): Promise<TrialBalanceEntry[]> {
    const db = SQLiteConfig.getDatabase();
    
    const entries = await (db as any).allAsync(
      'SELECT * FROM trial_balance_entries WHERE trial_balance_set_id = ? ORDER BY account_code',
      [parseInt(trialBalanceSetId)]
    );

    return entries.map((entry: any) => ({
      id: entry.id.toString(),
      trialBalanceSetId: entry.trial_balance_set_id.toString(),
      accountCode: entry.account_code,
      accountName: entry.account_name,
      previousBalance: entry.previous_balance || 0,
      currentBalance: entry.current_balance || 0,
      debitAmount: entry.debit_amount,
      creditAmount: entry.credit_amount,
      calculatedBalance: entry.balance,
      accountCategory: entry.account_type
    })) as TrialBalanceEntry[];
  }

  // Generated Statements operations
  async saveGeneratedStatement(statement: Omit<GeneratedStatement, 'id' | 'generatedAt'>): Promise<GeneratedStatement> {
    const db = SQLiteConfig.getDatabase();
    
    const result = await (db as any).runAsync(
      `INSERT INTO generated_statements 
       (company_id, trial_balance_set_id, statement_type, file_name, file_path, excel_data, formatting, notes)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [parseInt(statement.trialBalanceSetId), parseInt(statement.trialBalanceSetId), statement.statementType,
       'statement.xlsx', '', statement.excelBuffer || null, 
       JSON.stringify({
         fontName: statement.fontName,
         fontSize: statement.fontSize,
         columnWidths: statement.columnWidths,
         hasGreenBackground: statement.hasGreenBackground
       }), '']
    );

    const savedStatement = await (db as any).getAsync(
      'SELECT * FROM generated_statements WHERE id = ?',
      [result.lastID]
    );

    const formatting = savedStatement.formatting ? JSON.parse(savedStatement.formatting) : {};

    return {
      id: savedStatement.id.toString(),
      trialBalanceSetId: savedStatement.trial_balance_set_id.toString(),
      statementType: savedStatement.statement_type,
      data: statement.data,
      excelBuffer: savedStatement.excel_data,
      generatedAt: new Date(savedStatement.generated_at),
      fontName: formatting.fontName || 'Arial',
      fontSize: formatting.fontSize || 12,
      columnWidths: formatting.columnWidths || [],
      hasGreenBackground: formatting.hasGreenBackground || false
    } as GeneratedStatement;
  }

  async getGeneratedStatements(companyId: string): Promise<GeneratedStatement[]> {
    const db = SQLiteConfig.getDatabase();
    
    const statements = await (db as any).allAsync(
      `SELECT gs.*, tbs.file_name as trial_balance_file_name, c.name as company_name
       FROM generated_statements gs
       JOIN trial_balance_sets tbs ON gs.trial_balance_set_id = tbs.id
       JOIN companies c ON gs.company_id = c.id
       WHERE c.id = ?
       ORDER BY gs.generated_at DESC`,
      [parseInt(companyId)]
    );

    return statements.map((stmt: any) => {
      const formatting = stmt.formatting ? JSON.parse(stmt.formatting) : {};
      
      return {
        id: stmt.id.toString(),
        trialBalanceSetId: stmt.trial_balance_set_id.toString(),
        statementType: stmt.statement_type,
        data: null, // Statement data would need to be stored separately
        excelBuffer: stmt.excel_data,
        generatedAt: new Date(stmt.generated_at),
        fontName: formatting.fontName || 'Arial',
        fontSize: formatting.fontSize || 12,
        columnWidths: formatting.columnWidths || [],
        hasGreenBackground: formatting.hasGreenBackground || false
      };
    }) as GeneratedStatement[];
  }

  // Utility methods
  private generateFileHash(fileName: string, entries: any[]): string {
    const content = JSON.stringify({ fileName, entries });
    return createHash('sha256').update(content).digest('hex');
  }

  async getCompanyHistory(companyId: string): Promise<{
    company: Company;
    trialBalanceSets: TrialBalanceSet[];
    statements: GeneratedStatement[];
  }> {
    const [company, trialBalanceSets, statements] = await Promise.all([
      this.getCompany(companyId),
      this.getTrialBalanceSets(companyId),
      this.getGeneratedStatements(companyId)
    ]);

    if (!company) {
      throw new Error('Company not found');
    }

    return {
      company,
      trialBalanceSets,
      statements
    };
  }
}
