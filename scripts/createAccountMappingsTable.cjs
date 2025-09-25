const sqlite3 = require('sqlite3').verbose();
const path = require('path');

// Database file path
const dbPath = path.join(__dirname, '..', 'financial_statements.db');

// Default account mappings for new companies
const DEFAULT_MAPPINGS = [
  {
    noteType: 'cash',
    noteNumber: 7,
    noteTitle: 'เงินสดและรายการเทียบเท่าเงินสด',
    accountRanges: JSON.stringify({
      ranges: [{ from: 1000, to: 1099 }],
    })
  },
  {
    noteType: 'receivables',
    noteNumber: 8,
    noteTitle: 'ลูกหนี้การค้าและลูกหนี้อื่น',
    accountRanges: JSON.stringify({
      ranges: [{ from: 1140, to: 1215 }]
    })
  },
  {
    noteType: 'inventory',
    noteNumber: 9,
    noteTitle: 'สินค้าคงเหลือ',
    accountRanges: JSON.stringify({
      includes: [1510]
    })
  },
  {
    noteType: 'prepaid_expenses',
    noteNumber: 10,
    noteTitle: 'ค่าใช้จ่ายจ่ายล่วงหน้า',
    accountRanges: JSON.stringify({
      ranges: [{ from: 1400, to: 1439 }]
    })
  },
  {
    noteType: 'ppe_cost',
    noteNumber: 11,
    noteTitle: 'ที่ดิน อาคาร และอุปกรณ์',
    accountRanges: JSON.stringify({
      ranges: [{ from: 1610, to: 1659 }]
    })
  },
  {
    noteType: 'other_assets',
    noteNumber: 12,
    noteTitle: 'สินทรัพย์อื่น',
    accountRanges: JSON.stringify({
      ranges: [{ from: 1660, to: 1700 }]
    })
  },
  {
    noteType: 'bank_overdrafts',
    noteNumber: 15,
    noteTitle: 'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน',
    accountRanges: JSON.stringify({
      ranges: [{ from: 2001, to: 2009 }]
    })
  },
  {
    noteType: 'payables',
    noteNumber: 16,
    noteTitle: 'เจ้าหนี้การค้าและเจ้าหนี้อื่น',
    accountRanges: JSON.stringify({
      ranges: [{ from: 2010, to: 2999 }],
      excludes: [2030, 2045, 2050, 2051, 2052, 2100, 2101, 2102, 2103, 2120, 2121, 2122, 2123]
    })
  },
  {
    noteType: 'short_term_loans',
    noteNumber: 17,
    noteTitle: 'เงินกู้ยืมระยะสั้น',
    accountRanges: JSON.stringify({
      includes: [2030]
    })
  },
  {
    noteType: 'income_tax_payable',
    noteNumber: 18,
    noteTitle: 'ภาษีเงินได้นิติบุคคลค้างจ่าย',
    accountRanges: JSON.stringify({
      includes: [2045]
    })
  },
  {
    noteType: 'long_term_loans_fi',
    noteNumber: 19,
    noteTitle: 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน',
    accountRanges: JSON.stringify({
      ranges: [{ from: 2120, to: 2123 }],
      excludes: [2121]
    })
  },
  {
    noteType: 'long_term_loans_other',
    noteNumber: 20,
    noteTitle: 'เงินกู้ยืมระยะยาวอื่น',
    accountRanges: JSON.stringify({
      includes: [2050, 2051, 2052, 2100, 2101, 2102, 2103]
    })
  }
];

// DEPRECATED: Use scripts/setupSQLiteDatabase.js to create and seed tables.
console.error('[DEPRECATED] createAccountMappingsTable.cjs has been superseded by setupSQLiteDatabase.js');
process.exit(0);

/*
function createAccountMappingsTable() {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database(dbPath, (err) => {
      if (err) {
        reject(err);
        return;
      }
      console.log('Connected to SQLite database for account mappings migration');
    });

    // Create the company_account_mappings table
    const createTableSQL = `
      CREATE TABLE IF NOT EXISTS company_account_mappings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        note_type VARCHAR(50) NOT NULL,
        note_number INTEGER DEFAULT NULL,
        note_title VARCHAR(255) DEFAULT NULL,
        account_ranges TEXT NOT NULL,
        is_active BOOLEAN DEFAULT 1,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE,
        UNIQUE(company_id, note_type)
      );
    `;

    db.run(createTableSQL, (err) => {
      if (err) {
        reject(err);
        return;
      }
      console.log('✅ company_account_mappings table created successfully');

      // Get all existing companies to populate default mappings
      db.all('SELECT id FROM companies', (err, companies) => {
        if (err) {
          reject(err);
          return;
        }

        console.log(`📋 Found ${companies.length} existing companies`);

        if (companies.length === 0) {
          console.log('✅ No existing companies to populate');
          db.close();
          resolve();
          return;
        }

        // Populate default mappings for all existing companies
        let completed = 0;
        const totalMappings = companies.length * DEFAULT_MAPPINGS.length;

        companies.forEach(company => {
          DEFAULT_MAPPINGS.forEach(mapping => {
            const insertSQL = `
              INSERT OR IGNORE INTO company_account_mappings 
              (company_id, note_type, note_number, note_title, account_ranges) 
              VALUES (?, ?, ?, ?, ?)
            `;

            db.run(insertSQL, [
              company.id,
              mapping.noteType,
              mapping.noteNumber,
              mapping.noteTitle,
              mapping.accountRanges
            ], (err) => {
              if (err) {
                console.error(`❌ Error inserting mapping for company ${company.id}, note ${mapping.noteType}:`, err);
              } else {
                console.log(`✅ Created ${mapping.noteType} mapping for company ${company.id}`);
              }

              completed++;
              if (completed === totalMappings) {
                console.log(`🎉 Successfully populated ${totalMappings} account mappings`);
                db.close();
                resolve();
              }
            });
          });
        });
      });
    });
  });
}

// Run the migration
if (require.main === module) {
  createAccountMappingsTable()
    .then(() => {
      console.log('🎯 Account mappings table migration completed successfully!');
      process.exit(0);
    })
    .catch((error) => {
      console.error('💥 Migration failed:', error);
      process.exit(1);
    });
}

module.exports = { createAccountMappingsTable, DEFAULT_MAPPINGS };
*/
