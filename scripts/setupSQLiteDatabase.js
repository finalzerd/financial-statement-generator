import Database from 'sqlite3';

async function setupDatabase() {
  console.log('🚀 Starting SQLite database setup...');
  
  const dbPath = './financial_statements.db';
  const db = new Database.Database(dbPath);
  
  // Promisify database methods
  const runAsync = (sql, params = []) => {
    return new Promise((resolve, reject) => {
      db.run(sql, params, function(err) {
        if (err) reject(err);
        else resolve({ lastID: this.lastID, changes: this.changes });
      });
    });
  };

  const getAsync = (sql, params = []) => {
    return new Promise((resolve, reject) => {
      db.get(sql, params, (err, row) => {
        if (err) reject(err);
        else resolve(row);
      });
    });
  };

  try {
    // Test connection
    await getAsync('SELECT 1');
    console.log('✅ Database connection successful');

    // Create tables
    await runAsync('BEGIN TRANSACTION');

    // Companies table
    await runAsync(`
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
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
      )
    `);

    // Trial Balance Sets table
    await runAsync(`
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
    await runAsync(`
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
    await runAsync(`
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
    await runAsync(`
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

    await runAsync('COMMIT');

    console.log('✅ Database setup completed successfully!');
    console.log('📊 Tables created:');
    console.log('   - companies');
    console.log('   - trial_balance_sets'); 
    console.log('   - trial_balance_entries');
    console.log('   - generated_statements');
    console.log('   - company_settings');
    console.log(`📍 Database file: ${dbPath}`);

    // Insert a sample company for testing
    const result = await runAsync(
      `INSERT INTO companies (name, company_type) VALUES (?, ?)`,
      ['ตัวอย่างบริษัท จำกัด', 'บริษัทจำกัด']
    );
    
    console.log('🎯 Sample company created with ID:', result.lastID);

  } catch (error) {
    await runAsync('ROLLBACK');
    console.error('❌ Database setup failed:', error);
    process.exit(1);
  } finally {
    db.close((err) => {
      if (err) {
        console.error('Error closing database:', err);
      } else {
        console.log('✅ Database connection closed');
      }
    });
  }
}

setupDatabase();
