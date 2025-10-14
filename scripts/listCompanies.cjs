#!/usr/bin/env node
const sqlite3 = require('sqlite3');

const db = new sqlite3.Database('./financial_statements.db');

const sql = 'SELECT id, name, thai_name, company_type FROM companies ORDER BY id';

db.all(sql, (err, rows) => {
  if (err) {
    console.error('❌ Failed to query companies:', err.message);
    process.exitCode = 1;
  } else if (!rows || rows.length === 0) {
    console.log('No companies found.');
  } else {
    console.log('Companies in database:');
    rows.forEach((row) => {
      console.log(`  - ID ${row.id}: ${row.name || row.thai_name || '(unnamed)'} [${row.company_type || 'unknown type'}]`);
    });
  }
  db.close();
});
