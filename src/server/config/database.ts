import { Pool } from 'pg';

export class DatabaseConfig {
  private static pool: Pool;

  static initialize(): Pool {
    this.pool = new Pool({
      host: process.env.DB_HOST || 'localhost',
      port: parseInt(process.env.DB_PORT || '5432'),
      database: process.env.DB_NAME || 'financial_statements',
      user: process.env.DB_USER || 'postgres',
      password: process.env.DB_PASSWORD || 'password',
      ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false,
      max: 20, // Maximum number of clients in the pool
      idleTimeoutMillis: 30000, // Close clients after 30 seconds of inactivity
      connectionTimeoutMillis: 2000, // Return an error after 2 seconds if connection cannot be established
    });

    console.log('✅ Database connection pool initialized');
    console.log(`📍 Connecting to: ${process.env.DB_HOST}:${process.env.DB_PORT}/${process.env.DB_NAME}`);
    
    return this.pool;
  }

  static getPool(): Pool {
    if (!this.pool) {
      throw new Error('Database not initialized. Call DatabaseConfig.initialize() first.');
    }
    return this.pool;
  }

  static async testConnection(): Promise<boolean> {
    try {
      const client = await this.pool.connect();
      await client.query('SELECT 1');
      client.release();
      console.log('✅ Database connection test successful');
      return true;
    } catch (error) {
      console.error('❌ Database connection test failed:', error);
      return false;
    }
  }

  static async createTables(): Promise<void> {
    const client = await this.pool.connect();
    
    try {
      console.log('🏗️  Creating database tables...');
      
      // Create companies table
      await client.query(`
        CREATE TABLE IF NOT EXISTS companies (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          name VARCHAR(255) NOT NULL,
          type VARCHAR(50) NOT NULL CHECK (type IN ('ห้างหุ้นส่วนจำกัด', 'บริษัทจำกัด')),
          registration_number VARCHAR(100),
          address TEXT,
          business_description TEXT,
          tax_id VARCHAR(50),
          default_reporting_year INTEGER NOT NULL DEFAULT EXTRACT(YEAR FROM CURRENT_DATE),
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
      `);
      console.log('📋 Companies table created');

      // Create trial_balance_sets table
      await client.query(`
        CREATE TABLE IF NOT EXISTS trial_balance_sets (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          company_id UUID REFERENCES companies(id) ON DELETE CASCADE,
          reporting_year INTEGER NOT NULL,
          reporting_period VARCHAR(50) NOT NULL,
          processing_type VARCHAR(20) NOT NULL CHECK (processing_type IN ('single-year', 'multi-year')),
          file_name VARCHAR(255) NOT NULL,
          file_hash VARCHAR(64) NOT NULL,
          uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          status VARCHAR(20) DEFAULT 'processed' CHECK (status IN ('processed', 'archived')),
          metadata JSONB DEFAULT '{}',
          
          UNIQUE(company_id, reporting_year, file_hash)
        );
      `);
      console.log('📊 Trial balance sets table created');

      // Create trial_balance_entries table  
      await client.query(`
        CREATE TABLE IF NOT EXISTS trial_balance_entries (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          trial_balance_set_id UUID REFERENCES trial_balance_sets(id) ON DELETE CASCADE,
          account_code VARCHAR(20) NOT NULL,
          account_name VARCHAR(255) NOT NULL,
          previous_balance DECIMAL(15,2) DEFAULT 0,
          current_balance DECIMAL(15,2) DEFAULT 0,
          debit_amount DECIMAL(15,2) DEFAULT 0,
          credit_amount DECIMAL(15,2) DEFAULT 0,
          calculated_balance DECIMAL(15,2) DEFAULT 0,
          account_category VARCHAR(20) NOT NULL CHECK (account_category IN ('asset', 'liability', 'equity', 'revenue', 'expense'))
        );
      `);
      console.log('📝 Trial balance entries table created');

      // Create generated_statements table
      await client.query(`
        CREATE TABLE IF NOT EXISTS generated_statements (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          trial_balance_set_id UUID REFERENCES trial_balance_sets(id) ON DELETE CASCADE,
          statement_type VARCHAR(50) NOT NULL,
          data JSONB NOT NULL,
          excel_buffer BYTEA,
          generated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          font_name VARCHAR(50) DEFAULT 'TH Sarabun New',
          font_size INTEGER DEFAULT 14,
          column_widths JSONB,
          has_green_background BOOLEAN DEFAULT true,
          
          UNIQUE(trial_balance_set_id, statement_type)
        );
      `);
      console.log('📄 Generated statements table created');

      // Create company_settings table for customization
      await client.query(`
        CREATE TABLE IF NOT EXISTS company_settings (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          company_id UUID REFERENCES companies(id) ON DELETE CASCADE,
          setting_key VARCHAR(100) NOT NULL,
          setting_value TEXT NOT NULL,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          
          UNIQUE(company_id, setting_key)
        );
      `);
      console.log('⚙️  Company settings table created');

      // Create indexes for performance
      await client.query(`
        CREATE INDEX IF NOT EXISTS idx_trial_balance_company_year ON trial_balance_sets(company_id, reporting_year);
        CREATE INDEX IF NOT EXISTS idx_entries_set_code ON trial_balance_entries(trial_balance_set_id, account_code);
        CREATE INDEX IF NOT EXISTS idx_statements_set_type ON generated_statements(trial_balance_set_id, statement_type);
        CREATE INDEX IF NOT EXISTS idx_company_settings ON company_settings(company_id, setting_key);
      `);
      console.log('🔍 Database indexes created');

      console.log('✅ All database tables created successfully');

    } catch (error) {
      console.error('❌ Error creating database tables:', error);
      throw error;
    } finally {
      client.release();
    }
  }

  static async close(): Promise<void> {
    if (this.pool) {
      await this.pool.end();
      console.log('🔌 Database connection pool closed');
    }
  }
}
