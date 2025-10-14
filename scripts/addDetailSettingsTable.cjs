#!/usr/bin/env node
import Database from 'sqlite3';

const DETAIL_SETTINGS_TABLE_SQL = `
  CREATE TABLE IF NOT EXISTS company_detail_settings (
    company_id INTEGER PRIMARY KEY,
    detail_one_mode TEXT DEFAULT 'auto',
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE
  )
`;

function ensureDetailSettingsTable() {
  console.log('🔧 Ensuring company_detail_settings table exists...');
  const db = new Database.Database('./financial_statements.db');

  db.run(DETAIL_SETTINGS_TABLE_SQL, (err) => {
    if (err) {
      console.error('❌ Failed to create table:', err.message);
      process.exitCode = 1;
    } else {
      console.log('✅ company_detail_settings table is ready.');
    }

    db.close((closeErr) => {
      if (closeErr) {
        console.error('Error closing database:', closeErr.message);
        process.exitCode = 1;
      } else {
        console.log('✅ Database connection closed.');
      }
    });
  });
}

ensureDetailSettingsTable();
