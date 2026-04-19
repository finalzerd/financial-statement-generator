const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./financial_statements.db');

console.log("\n=== Company 60 - PPE Mappings ===\n");

db.all(`
  SELECT id, company_id, note_type, note_number, note_title, account_ranges, is_active 
  FROM company_account_mappings 
  WHERE company_id = 60 AND note_type LIKE 'ppe%'
  ORDER BY note_type
`, [], (err, rows) => {
  if (err) {
    console.error('Error:', err);
    db.close();
    return;
  }

  if (rows.length === 0) {
    console.log("No PPE mappings found for company 60");
    db.close();
    return;
  }

  rows.forEach(row => {
    console.log(`\n--- ${row.note_type} (ID: ${row.id}) ---`);
    console.log(`Note Number: ${row.note_number}`);
    console.log(`Note Title: ${row.note_title}`);
    console.log(`Is Active: ${row.is_active}`);
    console.log(`Raw JSON: ${row.account_ranges}`);
    
    try {
      const rules = JSON.parse(row.account_ranges);
      console.log(`\nParsed Rules:`);
      console.log(`  Ranges:`, rules.ranges || []);
      console.log(`  Includes:`, rules.includes || []);
      console.log(`  Excludes:`, rules.excludes || []);
    } catch (e) {
      console.log(`  (Could not parse account_ranges)`);
    }
  });

  console.log(`\n\nTotal PPE mappings for company 60: ${rows.length}`);
  
  db.close();
});
