import Database from 'sqlite3';

async function readDatabase() {
  console.log('📖 Reading SQLite Database...\n');
  
  const db = new Database.Database('./financial_statements.db');
  
  const allAsync = (sql) => {
    return new Promise((resolve, reject) => {
      db.all(sql, (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });
  };

  try {
    // List all tables
    console.log('📋 Available Tables:');
    const tables = await allAsync(`
      SELECT name FROM sqlite_master 
      WHERE type='table' AND name NOT LIKE 'sqlite_%'
    `);
    tables.forEach(table => console.log(`  - ${table.name}`));
    console.log('');

    // Show companies
    console.log('🏢 Companies:');
    const companies = await allAsync('SELECT * FROM companies');
    if (companies.length > 0) {
      companies.forEach(company => {
        console.log(`  ID: ${company.id}`);
        console.log(`  Name: ${company.name}`);
        console.log(`  Type: ${company.company_type}`);
        console.log(`  Created: ${company.created_at}`);
        console.log('  ---');
      });
    } else {
      console.log('  No companies found');
    }
    console.log('');

    // Show trial balance sets
    console.log('📊 Trial Balance Sets:');
    const trialSets = await allAsync('SELECT * FROM trial_balance_sets');
    if (trialSets.length > 0) {
      trialSets.forEach(set => {
        console.log(`  ID: ${set.id}`);
        console.log(`  File: ${set.file_name}`);
        console.log(`  Company ID: ${set.company_id}`);
        console.log(`  Upload Date: ${set.upload_date}`);
        console.log('  ---');
      });
    } else {
      console.log('  No trial balance sets found');
    }
    console.log('');

    // Show database schema
    console.log('🏗️  Database Schema:');
    for (const table of tables) {
      console.log(`\n${table.name.toUpperCase()} TABLE:`);
      const schema = await allAsync(`PRAGMA table_info(${table.name})`);
      schema.forEach(col => {
        console.log(`  ${col.name}: ${col.type}${col.pk ? ' (PRIMARY KEY)' : ''}`);
      });
    }

    // Database file info
    console.log('\n📁 Database File Info:');
    console.log(`  Location: ./financial_statements.db`);
    console.log(`  Total Tables: ${tables.length}`);
    console.log(`  Total Companies: ${companies.length}`);
    console.log(`  Total Trial Balance Sets: ${trialSets.length}`);

  } catch (error) {
    console.error('❌ Error reading database:', error);
  } finally {
    db.close((err) => {
      if (err) {
        console.error('Error closing database:', err);
      } else {
        console.log('\n✅ Database connection closed');
      }
    });
  }
}

readDatabase();
