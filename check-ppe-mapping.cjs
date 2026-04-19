const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./financial_statements.db');

console.log("Checking PPE mappings in database...\n");

db.all(`SELECT id, company_id, note_type, note_number, note_title, account_ranges, is_active 
        FROM company_account_mappings 
        WHERE note_type LIKE 'ppe%' 
        ORDER BY company_id, note_type`, [], (err, rows) => {
  if (err) {
    console.error('Error:', err);
  } else {
    rows.forEach(row => {
      console.log(`\n--- Mapping ID: ${row.id} ---`);
      console.log(`Company ID: ${row.company_id}`);
      console.log(`Note Type: ${row.note_type}`);
      console.log(`Note Number: ${row.note_number}`);
      console.log(`Note Title: ${row.note_title}`);
      console.log(`Is Active: ${row.is_active}`);
      console.log(`Account Ranges: ${row.account_ranges}`);
      
      // Parse and display the ranges nicely
      try {
        const rules = JSON.parse(row.account_ranges);
        console.log(`  Ranges:`, rules.ranges || []);
        console.log(`  Includes:`, rules.includes || []);
        console.log(`  Excludes:`, rules.excludes || []);
      } catch (e) {
        console.log(`  (Could not parse account_ranges)`);
      }
    });
    
    console.log(`\n\nTotal PPE mappings found: ${rows.length}`);
  }
  
  db.close();
});
