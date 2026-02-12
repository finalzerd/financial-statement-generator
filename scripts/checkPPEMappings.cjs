const sqlite3 = require('sqlite3').verbose();
const path = require('path');

const dbPath = path.join(__dirname, '..', 'financial_statements.db');
const db = new sqlite3.Database(dbPath);

db.all(
  `SELECT note_type, note_title, account_ranges 
   FROM company_account_mappings 
   WHERE note_type IN ('ppe_cost', 'intangible_assets_cost') 
   ORDER BY note_type`,
  (err, rows) => {
    if (err) {
      console.error('Error:', err);
    } else {
      rows.forEach(row => {
        console.log('\n' + '='.repeat(60));
        console.log('Note Type:', row.note_type);
        console.log('Title:', row.note_title);
        console.log('Rules:', row.account_ranges);
        
        // Parse and display the rules
        try {
          const rules = JSON.parse(row.account_ranges);
          console.log('\nParsed Rules:');
          console.log('  Ranges:', JSON.stringify(rules.ranges, null, 2));
          console.log('  Includes:', rules.includes || 'none');
          console.log('  Excludes:', rules.excludes || 'none');
        } catch (e) {
          console.log('  (Could not parse JSON)');
        }
      });
    }
    db.close();
  }
);
