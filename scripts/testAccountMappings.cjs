const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./financial_statements.db');

db.all('SELECT note_type, note_number, note_title, account_ranges FROM company_account_mappings WHERE company_id = 1 LIMIT 5', (err, rows) => {
  if (err) {
    console.error(err);
  } else {
    console.log('Sample Account Mappings for Company 1:');
    console.log('='.repeat(80));
    rows.forEach(row => {
      console.log(`Note ${row.note_number}: ${row.note_title}`);
      console.log(`  Type: ${row.note_type}`);
      console.log(`  Rules: ${row.account_ranges}`);
      console.log('');
    });
  }
  db.close();
});
