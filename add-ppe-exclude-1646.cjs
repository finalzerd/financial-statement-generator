const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./financial_statements.db');
const readline = require('readline');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

console.log("\n=== Add 1646 to PPE Excludes ===\n");

// Step 1: List all companies
db.all(`SELECT id, name FROM companies ORDER BY id`, [], (err, companies) => {
  if (err) {
    console.error('Error fetching companies:', err);
    db.close();
    rl.close();
    return;
  }

  console.log("Available companies:");
  companies.forEach(c => {
    console.log(`  ${c.id}. ${c.name}`);
  });

  rl.question('\nEnter company ID (or "all" to update all companies): ', (answer) => {
    const companyIds = answer.toLowerCase() === 'all' 
      ? companies.map(c => c.id)
      : [parseInt(answer)];

    if (companyIds.some(id => isNaN(id))) {
      console.log('Invalid company ID');
      db.close();
      rl.close();
      return;
    }

    let updatedCount = 0;
    const totalUpdates = companyIds.length * 2; // Both ppe_cost and ppe_accum_depr

    companyIds.forEach(companyId => {
      // Update ppe_cost
      db.get(
        `SELECT account_ranges FROM company_account_mappings WHERE company_id = ? AND note_type = 'ppe_cost'`,
        [companyId],
        (err, row) => {
          if (err) {
            console.error(`Error fetching ppe_cost for company ${companyId}:`, err);
            updatedCount++;
            checkComplete();
            return;
          }

          if (!row) {
            console.log(`No ppe_cost mapping found for company ${companyId}`);
            updatedCount++;
            checkComplete();
            return;
          }

          try {
            const rules = JSON.parse(row.account_ranges);
            const excludes = rules.excludes || [];
            
            if (!excludes.includes(1646)) {
              excludes.push(1646);
              rules.excludes = excludes;

              db.run(
                `UPDATE company_account_mappings SET account_ranges = ?, updated_at = CURRENT_TIMESTAMP WHERE company_id = ? AND note_type = 'ppe_cost'`,
                [JSON.stringify(rules), companyId],
                (err) => {
                  if (err) {
                    console.error(`Error updating ppe_cost for company ${companyId}:`, err);
                  } else {
                    console.log(`✓ Added 1646 to ppe_cost excludes for company ${companyId}`);
                  }
                  updatedCount++;
                  checkComplete();
                }
              );
            } else {
              console.log(`  1646 already in ppe_cost excludes for company ${companyId}`);
              updatedCount++;
              checkComplete();
            }
          } catch (parseErr) {
            console.error(`Error parsing account_ranges for company ${companyId}:`, parseErr);
            updatedCount++;
            checkComplete();
          }
        }
      );

      // Update ppe_accum_depr (need to check if it exists first)
      db.get(
        `SELECT account_ranges FROM company_account_mappings WHERE company_id = ? AND note_type = 'ppe_accum_depr'`,
        [companyId],
        (err, row) => {
          if (err) {
            console.error(`Error fetching ppe_accum_depr for company ${companyId}:`, err);
            updatedCount++;
            checkComplete();
            return;
          }

          if (!row) {
            // Create ppe_accum_depr mapping if it doesn't exist
            const newRules = {
              ranges: [{ from: 1630, to: 1659 }],
              excludes: [1646]
            };

            db.run(
              `INSERT INTO company_account_mappings (company_id, note_type, note_number, note_title, account_ranges, is_active)
               VALUES (?, 'ppe_accum_depr', 11, 'ที่ดิน อาคาร และอุปกรณ์ (ค่าเสื่อมราคาสะสม)', ?, 1)`,
              [companyId, JSON.stringify(newRules)],
              (err) => {
                if (err) {
                  console.error(`Error creating ppe_accum_depr for company ${companyId}:`, err);
                } else {
                  console.log(`✓ Created ppe_accum_depr with 1646 in excludes for company ${companyId}`);
                }
                updatedCount++;
                checkComplete();
              }
            );
            return;
          }

          try {
            const rules = JSON.parse(row.account_ranges);
            const excludes = rules.excludes || [];
            
            if (!excludes.includes(1646)) {
              excludes.push(1646);
              rules.excludes = excludes;

              db.run(
                `UPDATE company_account_mappings SET account_ranges = ?, updated_at = CURRENT_TIMESTAMP WHERE company_id = ? AND note_type = 'ppe_accum_depr'`,
                [JSON.stringify(rules), companyId],
                (err) => {
                  if (err) {
                    console.error(`Error updating ppe_accum_depr for company ${companyId}:`, err);
                  } else {
                    console.log(`✓ Added 1646 to ppe_accum_depr excludes for company ${companyId}`);
                  }
                  updatedCount++;
                  checkComplete();
                }
              );
            } else {
              console.log(`  1646 already in ppe_accum_depr excludes for company ${companyId}`);
              updatedCount++;
              checkComplete();
            }
          } catch (parseErr) {
            console.error(`Error parsing account_ranges for company ${companyId}:`, parseErr);
            updatedCount++;
            checkComplete();
          }
        }
      );
    });

    function checkComplete() {
      if (updatedCount >= totalUpdates) {
        console.log('\n=== Update Complete ===');
        db.close();
        rl.close();
      }
    }
  });
});
