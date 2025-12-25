/**
 * Migration Script: Add Period Start/End Date Columns to Companies Table
 * 
 * This script adds:
 * - period_start_date TEXT DEFAULT '1 มกราคม'
 * - period_end_date TEXT DEFAULT '31 ธันวาคม'
 * 
 * Run with: node scripts/addPeriodDatesColumn.cjs
 */

const sqlite3 = require('sqlite3').verbose();
const path = require('path');

const DB_PATH = path.join(__dirname, '..', 'financial_statements.db');

function runMigration() {
  const db = new sqlite3.Database(DB_PATH, (err) => {
    if (err) {
      console.error('❌ Error opening database:', err.message);
      process.exit(1);
    }
    console.log('✅ Connected to financial_statements.db');
  });

  db.serialize(() => {
    console.log('\n🔄 Starting migration: Add period date columns...\n');

    // Check if columns already exist
    db.all("PRAGMA table_info(companies)", (err, columns) => {
      if (err) {
        console.error('❌ Error checking table structure:', err.message);
        db.close();
        process.exit(1);
      }

      const hasStartDate = columns.some(col => col.name === 'period_start_date');
      const hasEndDate = columns.some(col => col.name === 'period_end_date');

      if (hasStartDate && hasEndDate) {
        console.log('⚠️  Columns already exist. Migration skipped.');
        db.close();
        return;
      }

      // Add period_start_date column
      if (!hasStartDate) {
        db.run(
          `ALTER TABLE companies ADD COLUMN period_start_date TEXT DEFAULT '1 มกราคม'`,
          (err) => {
            if (err) {
              console.error('❌ Error adding period_start_date column:', err.message);
              db.close();
              process.exit(1);
            }
            console.log('✅ Added column: period_start_date');
          }
        );
      }

      // Add period_end_date column
      if (!hasEndDate) {
        db.run(
          `ALTER TABLE companies ADD COLUMN period_end_date TEXT DEFAULT '31 ธันวาคม'`,
          (err) => {
            if (err) {
              console.error('❌ Error adding period_end_date column:', err.message);
              db.close();
              process.exit(1);
            }
            console.log('✅ Added column: period_end_date');
          }
        );
      }

      // Update existing records to have default values
      db.run(
        `UPDATE companies 
         SET period_start_date = '1 มกราคม', 
             period_end_date = '31 ธันวาคม' 
         WHERE period_start_date IS NULL OR period_end_date IS NULL`,
        function(err) {
          if (err) {
            console.error('❌ Error updating existing records:', err.message);
          } else {
            console.log(`✅ Updated ${this.changes} existing records with default dates`);
          }

          // Verify migration
          db.all("SELECT id, name, period_start_date, period_end_date FROM companies", (err, rows) => {
            if (err) {
              console.error('❌ Error verifying migration:', err.message);
            } else {
              console.log('\n📊 Current companies with period dates:');
              if (rows.length === 0) {
                console.log('   (No companies in database)');
              } else {
                rows.forEach(row => {
                  console.log(`   ID ${row.id}: ${row.name} - ${row.period_start_date} to ${row.period_end_date}`);
                });
              }
            }

            console.log('\n✅ Migration completed successfully!\n');
            db.close();
          });
        }
      );
    });
  });
}

// Rollback function (optional, for reference)
function rollbackMigration() {
  console.log('⚠️  SQLite does not support DROP COLUMN directly.');
  console.log('To rollback, you would need to:');
  console.log('1. Create a new table without these columns');
  console.log('2. Copy data from old table');
  console.log('3. Drop old table');
  console.log('4. Rename new table');
  console.log('\nFor safety, rollback is not automated. Backup your database first!');
}

// Run migration
if (require.main === module) {
  const args = process.argv.slice(2);
  if (args.includes('--rollback')) {
    rollbackMigration();
  } else {
    runMigration();
  }
}

module.exports = { runMigration, rollbackMigration };
