import { Pool } from 'pg';
import { DatabaseConfig } from '../config/database';
import type { Company, TrialBalanceSet, TrialBalanceEntry, GeneratedStatement, CSVTrialBalanceEntry } from '../../types/database';
import * as crypto from 'crypto';

export class DatabaseService {
  private static get pool(): Pool {
    return DatabaseConfig.getPool();
  }

  // ============== COMPANY OPERATIONS ==============
  
  static async createCompany(companyData: Omit<Company, 'id' | 'createdAt' | 'updatedAt'>): Promise<Company> {
    const query = `
      INSERT INTO companies (name, type, registration_number, address, business_description, tax_id, default_reporting_year)
      VALUES ($1, $2, $3, $4, $5, $6, $7)
      RETURNING *
    `;

    const result = await this.pool.query(query, [
      companyData.name,
      companyData.type,
      companyData.registrationNumber,
      companyData.address,
      companyData.businessDescription,
      companyData.taxId,
      companyData.defaultReportingYear
    ]);

    return this.mapCompanyRow(result.rows[0]);
  }

  static async getCompanies(): Promise<Company[]> {
    const query = 'SELECT * FROM companies ORDER BY created_at DESC';
    const result = await this.pool.query(query);
    return result.rows.map(this.mapCompanyRow);
  }

  static async getCompanyById(id: string): Promise<Company | null> {
    const query = 'SELECT * FROM companies WHERE id = $1';
    const result = await this.pool.query(query, [id]);
    return result.rows.length > 0 ? this.mapCompanyRow(result.rows[0]) : null;
  }

  static async updateCompany(id: string, updates: Partial<Company>): Promise<Company | null> {
    const setClause = [];
    const values = [];
    let paramIndex = 1;

    if (updates.name) {
      setClause.push(`name = $${paramIndex++}`);
      values.push(updates.name);
    }
    if (updates.type) {
      setClause.push(`type = $${paramIndex++}`);
      values.push(updates.type);
    }
    if (updates.registrationNumber !== undefined) {
      setClause.push(`registration_number = $${paramIndex++}`);
      values.push(updates.registrationNumber);
    }
    if (updates.address !== undefined) {
      setClause.push(`address = $${paramIndex++}`);
      values.push(updates.address);
    }
    if (updates.businessDescription !== undefined) {
      setClause.push(`business_description = $${paramIndex++}`);
      values.push(updates.businessDescription);
    }
    if (updates.taxId !== undefined) {
      setClause.push(`tax_id = $${paramIndex++}`);
      values.push(updates.taxId);
    }
    if (updates.defaultReportingYear) {
      setClause.push(`default_reporting_year = $${paramIndex++}`);
      values.push(updates.defaultReportingYear);
    }

    if (setClause.length === 0) {
      return this.getCompanyById(id);
    }

    setClause.push(`updated_at = CURRENT_TIMESTAMP`);
    values.push(id);

    const query = `
      UPDATE companies 
      SET ${setClause.join(', ')} 
      WHERE id = $${paramIndex}
      RETURNING *
    `;

    const result = await this.pool.query(query, values);
    return result.rows.length > 0 ? this.mapCompanyRow(result.rows[0]) : null;
  }

  // ============== TRIAL BALANCE OPERATIONS ==============

  static async saveTrialBalanceSet(
    companyId: string,
    csvEntries: CSVTrialBalanceEntry[],
    metadata: {
      fileName: string;
      processingType: 'single-year' | 'multi-year';
      reportingYear: number;
      reportingPeriod: string;
    }
  ): Promise<TrialBalanceSet> {
    const client = await this.pool.connect();

    try {
      await client.query('BEGIN');

      // Generate file hash for duplicate detection
      const fileHash = this.generateFileHash(csvEntries);

      // Check for duplicates
      const duplicateCheck = await client.query(
        'SELECT id FROM trial_balance_sets WHERE company_id = $1 AND reporting_year = $2 AND file_hash = $3',
        [companyId, metadata.reportingYear, fileHash]
      );

      if (duplicateCheck.rows.length > 0) {
        throw new Error('Duplicate trial balance data detected for this company and year');
      }

      // Create trial balance set
      const setQuery = `
        INSERT INTO trial_balance_sets (
          company_id, reporting_year, reporting_period, processing_type,
          file_name, file_hash, metadata
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7)
        RETURNING *
      `;

      const setMetadata = {
        totalEntries: csvEntries.length,
        hasInventory: csvEntries.some(e => e.รหัสบัญชี === '1510'),
        hasPPE: csvEntries.some(e => e.รหัสบัญชี.startsWith('161')),
        accountCodeRange: {
          min: Math.min(...csvEntries.map(e => parseInt(e.รหัสบัญชี) || 0)).toString(),
          max: Math.max(...csvEntries.map(e => parseInt(e.รหัสบัญชี) || 0)).toString()
        }
      };

      const setResult = await client.query(setQuery, [
        companyId,
        metadata.reportingYear,
        metadata.reportingPeriod,
        metadata.processingType,
        metadata.fileName,
        fileHash,
        JSON.stringify(setMetadata)
      ]);

      const trialBalanceSet = this.mapTrialBalanceSetRow(setResult.rows[0]);

      // Insert trial balance entries
      for (const csvEntry of csvEntries) {
        const entryQuery = `
          INSERT INTO trial_balance_entries (
            trial_balance_set_id, account_code, account_name,
            previous_balance, current_balance, debit_amount, credit_amount,
            calculated_balance, account_category
          )
          VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
        `;

        await client.query(entryQuery, [
          trialBalanceSet.id,
          csvEntry.รหัสบัญชี,
          csvEntry.ชื่อบัญชี,
          csvEntry.ยอดยกมาต้นงวด || 0,
          csvEntry.ยอดยกมางวดนี้ || 0,
          csvEntry.เดบิต || 0,
          csvEntry.เครดิต || 0,
          this.calculateBalance(csvEntry),
          this.categorizeAccount(csvEntry.รหัสบัญชี)
        ]);
      }

      await client.query('COMMIT');
      console.log(`✅ Saved trial balance set with ${csvEntries.length} entries`);
      return trialBalanceSet;

    } catch (error) {
      await client.query('ROLLBACK');
      console.error('❌ Error saving trial balance set:', error);
      throw error;
    } finally {
      client.release();
    }
  }

  static async getTrialBalanceHistory(companyId: string): Promise<TrialBalanceSet[]> {
    const query = `
      SELECT * FROM trial_balance_sets 
      WHERE company_id = $1 
      ORDER BY reporting_year DESC, uploaded_at DESC
    `;
    const result = await this.pool.query(query, [companyId]);
    return result.rows.map(this.mapTrialBalanceSetRow);
  }

  static async getTrialBalanceEntries(trialBalanceSetId: string): Promise<TrialBalanceEntry[]> {
    const query = `
      SELECT * FROM trial_balance_entries 
      WHERE trial_balance_set_id = $1 
      ORDER BY account_code
    `;
    const result = await this.pool.query(query, [trialBalanceSetId]);
    return result.rows.map(this.mapTrialBalanceEntryRow);
  }

  static async getTrialBalanceSetById(id: string): Promise<TrialBalanceSet | null> {
    const query = 'SELECT * FROM trial_balance_sets WHERE id = $1';
    const result = await this.pool.query(query, [id]);
    return result.rows.length > 0 ? this.mapTrialBalanceSetRow(result.rows[0]) : null;
  }

  // ============== GENERATED STATEMENTS OPERATIONS ==============

  static async saveGeneratedStatement(
    trialBalanceSetId: string,
    statementType: string,
    data: any,
    excelBuffer?: Buffer
  ): Promise<GeneratedStatement> {
    const query = `
      INSERT INTO generated_statements (
        trial_balance_set_id, statement_type, data, excel_buffer,
        font_name, font_size, column_widths, has_green_background
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
      ON CONFLICT (trial_balance_set_id, statement_type)
      DO UPDATE SET 
        data = EXCLUDED.data,
        excel_buffer = EXCLUDED.excel_buffer,
        generated_at = CURRENT_TIMESTAMP
      RETURNING *
    `;

    const columnWidths = this.getColumnWidthsForStatementType(statementType);
    const hasGreenBackground = statementType !== 'balance-sheet-assets';

    const result = await this.pool.query(query, [
      trialBalanceSetId,
      statementType,
      JSON.stringify(data),
      excelBuffer,
      'TH Sarabun New',
      14,
      JSON.stringify(columnWidths),
      hasGreenBackground
    ]);

    console.log(`✅ Saved generated statement: ${statementType}`);
    return this.mapGeneratedStatementRow(result.rows[0]);
  }

  static async getGeneratedStatements(trialBalanceSetId: string): Promise<GeneratedStatement[]> {
    const query = `
      SELECT * FROM generated_statements 
      WHERE trial_balance_set_id = $1 
      ORDER BY generated_at DESC
    `;
    const result = await this.pool.query(query, [trialBalanceSetId]);
    return result.rows.map(this.mapGeneratedStatementRow);
  }

  static async getStatementById(id: string): Promise<GeneratedStatement | null> {
    const query = 'SELECT * FROM generated_statements WHERE id = $1';
    const result = await this.pool.query(query, [id]);
    return result.rows.length > 0 ? this.mapGeneratedStatementRow(result.rows[0]) : null;
  }

  // ============== HELPER METHODS ==============

  private static generateFileHash(csvEntries: CSVTrialBalanceEntry[]): string {
    const content = csvEntries
      .map(e => `${e.รหัสบัญชี}:${e.เดบิต}:${e.เครดิต}`)
      .join('|');
    return crypto.createHash('sha256').update(content).digest('hex');
  }

  private static calculateBalance(csvEntry: CSVTrialBalanceEntry): number {
    // Use your existing balance calculation logic
    const debit = csvEntry.เดบิต || 0;
    const credit = csvEntry.เครดิต || 0;
    const accountCode = csvEntry.รหัสบัญชี;
    
    // For assets and expenses (codes 1xxx and 5xxx), debit increases balance
    if (accountCode.startsWith('1') || accountCode.startsWith('5')) {
      return debit - credit;
    }
    // For liabilities, equity, and revenue (codes 2xxx, 3xxx, 4xxx), credit increases balance
    else {
      return credit - debit;
    }
  }

  private static categorizeAccount(accountCode: string): string {
    if (accountCode.startsWith('1')) return 'asset';
    if (accountCode.startsWith('2')) return 'liability';
    if (accountCode.startsWith('3')) return 'equity';
    if (accountCode.startsWith('4')) return 'revenue';
    if (accountCode.startsWith('5')) return 'expense';
    return 'asset';
  }

  private static getColumnWidthsForStatementType(statementType: string): number[] {
    // Based on ExcelJSFormatter column widths
    if (statementType === 'balance-sheet-assets') {
      return [3.94, 5.51, 6.30, 5.51, 22.05, 5.51, 11.02, 1.57, 11.02];
    } else if (statementType === 'equity-changes') {
      return [30, 2, 14, 2, 2, 14, 2, 2, 14];
    } else {
      return [3, 45, 3, 12, 3, 18, 18];
    }
  }

  // ============== ROW MAPPING METHODS ==============

  private static mapCompanyRow(row: any): Company {
    return {
      id: row.id,
      name: row.name,
      type: row.type,
      registrationNumber: row.registration_number,
      address: row.address,
      businessDescription: row.business_description,
      taxId: row.tax_id,
      defaultReportingYear: row.default_reporting_year,
      createdAt: new Date(row.created_at),
      updatedAt: new Date(row.updated_at)
    };
  }

  private static mapTrialBalanceSetRow(row: any): TrialBalanceSet {
    return {
      id: row.id,
      companyId: row.company_id,
      reportingYear: row.reporting_year,
      reportingPeriod: row.reporting_period,
      processingType: row.processing_type,
      fileName: row.file_name,
      fileHash: row.file_hash,
      uploadedAt: new Date(row.uploaded_at),
      status: row.status,
      metadata: row.metadata
    };
  }

  private static mapTrialBalanceEntryRow(row: any): TrialBalanceEntry {
    return {
      id: row.id,
      trialBalanceSetId: row.trial_balance_set_id,
      accountCode: row.account_code,
      accountName: row.account_name,
      previousBalance: parseFloat(row.previous_balance),
      currentBalance: parseFloat(row.current_balance),
      debitAmount: parseFloat(row.debit_amount),
      creditAmount: parseFloat(row.credit_amount),
      calculatedBalance: parseFloat(row.calculated_balance),
      accountCategory: row.account_category
    };
  }

  private static mapGeneratedStatementRow(row: any): GeneratedStatement {
    return {
      id: row.id,
      trialBalanceSetId: row.trial_balance_set_id,
      statementType: row.statement_type,
      data: row.data,
      excelBuffer: row.excel_buffer,
      generatedAt: new Date(row.generated_at),
      fontName: row.font_name,
      fontSize: row.font_size,
      columnWidths: row.column_widths || [],
      hasGreenBackground: row.has_green_background
    };
  }
}
