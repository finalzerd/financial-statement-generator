const path = require('path');
const sqlite3 = require('sqlite3').verbose();

(async function main(){
  const dbPath = path.resolve(__dirname, '..', 'financial_statements.db');
  console.log('Database file:', dbPath);
  const db = new sqlite3.Database(dbPath);
  const all = (sql, params=[]) => new Promise((res, rej)=> db.all(sql, params, (e, r)=> e?rej(e):res(r)));
  try {
    const tables = await all("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'");
    const has = tables.some(t => t.name === 'company_account_mappings');
    if(!has){
      console.log('Table company_account_mappings not found. Available tables:', tables.map(t=>t.name).join(', '));
    } else {
      const cols = await all('PRAGMA table_info(company_account_mappings)');
      console.log('\ncompany_account_mappings columns:');
      cols.forEach(c => console.log(` - ${c.name}: ${c.type}${c.pk? ' (PRIMARY KEY)': ''}`));
      const sample = await all('SELECT id, company_id, note_type, note_number, note_title, is_active FROM company_account_mappings ORDER BY company_id, note_number LIMIT 5');
      console.log('\nSample rows:');
      sample.forEach(r => console.log(r));
    }
  } catch (e) {
    console.error('Error:', e);
  } finally {
    db.close();
  }
})();
