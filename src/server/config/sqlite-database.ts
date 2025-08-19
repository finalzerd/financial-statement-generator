import Database from 'sqlite3';
import { promisify } from 'util';

export class SQLiteConfig {
  private static db: Database.Database;
  private static dbPath: string;

  static initialize(): Database.Database {
    this.dbPath = process.env.DB_PATH || './financial_statements.db';
    this.db = new Database.Database(this.dbPath);
    
    // Promisify common methods for easier async/await usage
    (this.db as any).runAsync = promisify(this.db.run.bind(this.db));
    (this.db as any).getAsync = promisify(this.db.get.bind(this.db));
    (this.db as any).allAsync = promisify(this.db.all.bind(this.db));

    console.log('✅ SQLite database initialized');
    console.log(`📍 Database file: ${this.dbPath}`);
    
    return this.db;
  }

  static getDatabase(): Database.Database {
    if (!this.db) {
      throw new Error('Database not initialized. Call SQLiteConfig.initialize() first.');
    }
    return this.db;
  }

  static async testConnection(): Promise<boolean> {
    try {
      await (this.db as any).getAsync('SELECT 1');
      console.log('✅ Database connection test successful');
      return true;
    } catch (error) {
      console.error('❌ Database connection failed:', error);
      return false;
    }
  }

  static async createTables(): Promise<void> {
    try {
      await (this.db as any).runAsync('BEGIN TRANSACTION');

      // Companies table
      await (this.db as any).runAsync(`
        CREATE TABLE IF NOT EXISTS companies (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL,
          thai_name TEXT,
          registration_number TEXT,
          company_type TEXT NOT NULL,
          address TEXT,
          phone TEXT,
          email TEXT,
          business_type TEXT,
          number_of_shares INTEGER,
          share_value REAL,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
      `);

      // Add share fields to existing companies table if they don't exist
      try {
        await (this.db as any).runAsync('ALTER TABLE companies ADD COLUMN number_of_shares INTEGER');
      } catch (e) {
        // Column already exists, ignore error
      }
      
      try {
        await (this.db as any).runAsync('ALTER TABLE companies ADD COLUMN share_value REAL');
      } catch (e) {
        // Column already exists, ignore error
      }

      // Trial Balance Sets table
      await (this.db as any).runAsync(`
        CREATE TABLE IF NOT EXISTS trial_balance_sets (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          company_id INTEGER REFERENCES companies(id) ON DELETE CASCADE,
          file_name TEXT NOT NULL,
          file_hash TEXT UNIQUE NOT NULL,
          upload_date DATETIME DEFAULT CURRENT_TIMESTAMP,
          period_start DATE,
          period_end DATE,
          currency TEXT DEFAULT 'THB',
          notes TEXT
        )
      `);

      // Trial Balance Entries table
      await (this.db as any).runAsync(`
        CREATE TABLE IF NOT EXISTS trial_balance_entries (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          trial_balance_set_id INTEGER REFERENCES trial_balance_sets(id) ON DELETE CASCADE,
          account_code TEXT NOT NULL,
          account_name TEXT NOT NULL,
          debit_amount REAL DEFAULT 0,
          credit_amount REAL DEFAULT 0,
          balance REAL DEFAULT 0,
          account_type TEXT,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
      `);

      // Generated Statements table
      await (this.db as any).runAsync(`
        CREATE TABLE IF NOT EXISTS generated_statements (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          company_id INTEGER REFERENCES companies(id) ON DELETE CASCADE,
          trial_balance_set_id INTEGER REFERENCES trial_balance_sets(id) ON DELETE CASCADE,
          statement_type TEXT NOT NULL,
          file_name TEXT NOT NULL,
          file_path TEXT,
          excel_data BLOB,
          formatting TEXT,
          generated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          notes TEXT
        )
      `);

      // Company Settings table
      await (this.db as any).runAsync(`
        CREATE TABLE IF NOT EXISTS company_settings (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          company_id INTEGER REFERENCES companies(id) ON DELETE CASCADE,
          setting_key TEXT NOT NULL,
          setting_value TEXT,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          UNIQUE(company_id, setting_key)
        )
      `);

      await (this.db as any).runAsync('COMMIT');
      console.log('✅ All database tables created successfully');
    } catch (error) {
      await (this.db as any).runAsync('ROLLBACK');
      console.error('❌ Error creating tables:', error);
      throw error;
    }
  }

  static async close(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      this.db.close((err) => {
        if (err) {
          reject(err);
        } else {
          console.log('✅ Database connection closed');
          resolve();
        }
      });
    });
  }
}
