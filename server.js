import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import sqlite3 from 'sqlite3';
// Uncomment next line to use PostgreSQL instead of SQLite
// import { DatabaseConfig } from './src/server/config/database.js';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3001;

// Database connection
const dbPath = './financial_statements.db';
const db = new sqlite3.Database(dbPath);

// Add share columns to companies table if they don't exist
db.run('ALTER TABLE companies ADD COLUMN number_of_shares INTEGER', (err) => {
  // Ignore error if column already exists
});

db.run('ALTER TABLE companies ADD COLUMN share_value REAL', (err) => {
  // Ignore error if column already exists
});

// Helper function to convert database row to Company format
const mapDbRowToCompany = (row) => ({
  id: row.id.toString(),
  name: row.name || row.thai_name || 'Unknown Company',
  type: row.company_type || 'บริษัทจำกัด',
  registrationNumber: row.registration_number,
  address: row.address,
  businessDescription: row.business_type,
  taxId: row.tax_id,
  numberOfShares: row.number_of_shares,
  shareValue: row.share_value,
  defaultReportingYear: new Date().getFullYear(), // Default to current year
  createdAt: new Date(row.created_at),
  updatedAt: new Date(row.updated_at || row.created_at)
});

// Middleware
app.use(cors({
  origin: process.env.FRONTEND_URL || 'http://localhost:5173',
  credentials: true
}));
app.use(express.json());

// Health check route
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    message: 'Financial Statement Generator API is running',
    database: 'SQLite',
    timestamp: new Date().toISOString()
  });
});

// Get all companies
app.get('/api/companies', (req, res) => {
  db.all('SELECT * FROM companies ORDER BY created_at DESC', (err, rows) => {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Database error',
        details: err.message 
      });
      return;
    }

    const companies = rows.map(mapDbRowToCompany);
    res.json({ 
      success: true, 
      data: companies,
      count: companies.length,
      message: companies.length > 0 ? 'Companies loaded successfully' : 'No companies found'
    });
  });
});

// Create new company
app.post('/api/companies', (req, res) => {
  const { 
    name, 
    type, 
    registrationNumber, 
    address, 
    businessDescription, 
    taxId, 
    defaultReportingYear,
    numberOfShares,
    shareValue
  } = req.body;
  
  if (!name || !type) {
    res.status(400).json({ 
      success: false, 
      error: 'Missing required fields',
      details: 'Name and type are required' 
    });
    return;
  }

  const now = new Date().toISOString();
  const query = `
    INSERT INTO companies (name, thai_name, company_type, registration_number, address, business_type, phone, email, number_of_shares, share_value, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?, ?)
  `;

  db.run(query, [
    name, 
    name, 
    type, 
    registrationNumber, 
    address, 
    businessDescription, 
    type === 'บริษัทจำกัด' ? numberOfShares : null,
    type === 'บริษัทจำกัด' ? shareValue : null,
    now, 
    now
  ], function(err) {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to create company',
        details: err.message 
      });
      return;
    }

    // Get the newly created company
    db.get('SELECT * FROM companies WHERE id = ?', [this.lastID], (err, row) => {
      if (err) {
        console.error('Database error:', err);
        res.status(500).json({ 
          success: false, 
          error: 'Failed to retrieve created company',
          details: err.message 
        });
        return;
      }

      const company = mapDbRowToCompany(row);
      res.status(201).json({ 
        success: true, 
        company,
        message: 'Company created successfully' 
      });
    });
  });
});

// Update existing company
app.put('/api/companies/:id', (req, res) => {
  const companyId = req.params.id;
  const { 
    name, 
    type, 
    registrationNumber, 
    address, 
    businessDescription, 
    taxId,
    numberOfShares,
    shareValue
  } = req.body;
  
  if (!name || !type) {
    res.status(400).json({ 
      success: false, 
      error: 'Missing required fields',
      details: 'Name and type are required' 
    });
    return;
  }

  const now = new Date().toISOString();
  const query = `
    UPDATE companies 
    SET name = ?, thai_name = ?, company_type = ?, registration_number = ?, 
        address = ?, business_type = ?, number_of_shares = ?, share_value = ?, updated_at = ?
    WHERE id = ?
  `;

  db.run(query, [
    name, 
    name, 
    type, 
    registrationNumber, 
    address, 
    businessDescription, 
    type === 'บริษัทจำกัด' ? numberOfShares : null,
    type === 'บริษัทจำกัด' ? shareValue : null,
    now, 
    companyId
  ], function(err) {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to update company',
        details: err.message 
      });
      return;
    }

    if (this.changes === 0) {
      res.status(404).json({ 
        success: false, 
        error: 'Company not found',
        details: `No company found with ID ${companyId}` 
      });
      return;
    }

    // Get the updated company
    db.get('SELECT * FROM companies WHERE id = ?', [companyId], (err, row) => {
      if (err) {
        console.error('Database error:', err);
        res.status(500).json({ 
          success: false, 
          error: 'Failed to retrieve updated company',
          details: err.message 
        });
        return;
      }

      const company = mapDbRowToCompany(row);
      res.json({ 
        success: true, 
        company,
        message: 'Company updated successfully' 
      });
    });
  });
});

// Delete existing company
app.delete('/api/companies/:id', (req, res) => {
  const companyId = req.params.id;
  
  // First check if company exists
  db.get('SELECT * FROM companies WHERE id = ?', [companyId], (err, company) => {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to check company existence',
        details: err.message 
      });
      return;
    }

    if (!company) {
      res.status(404).json({ 
        success: false, 
        error: 'Company not found',
        details: `No company found with ID ${companyId}` 
      });
      return;
    }

    // Check if company has any trial balance data
    db.get('SELECT COUNT(*) as count FROM trial_balance_sets WHERE company_id = ?', [companyId], (err, result) => {
      if (err) {
        console.error('Database error:', err);
        res.status(500).json({ 
          success: false, 
          error: 'Failed to check trial balance data',
          details: err.message 
        });
        return;
      }

      const hasTrialBalanceData = result.count > 0;

      if (hasTrialBalanceData) {
        res.status(400).json({ 
          success: false, 
          error: 'Cannot delete company',
          details: 'This company has trial balance data. Please delete all trial balance records first.' 
        });
        return;
      }

      // Safe to delete company
      db.run('DELETE FROM companies WHERE id = ?', [companyId], function(err) {
        if (err) {
          console.error('Database error:', err);
          res.status(500).json({ 
            success: false, 
            error: 'Failed to delete company',
            details: err.message 
          });
          return;
        }

        if (this.changes === 0) {
          res.status(404).json({ 
            success: false, 
            error: 'Company not found',
            details: `No company found with ID ${companyId}` 
          });
          return;
        }

        res.json({ 
          success: true, 
          message: `Company "${company.name}" deleted successfully` 
        });
      });
    });
  });
});

// ============== TRIAL BALANCE ENDPOINTS ==============

// Save trial balance data
app.post('/api/trial-balance', (req, res) => {
  const { companyId, trialBalanceEntries, metadata } = req.body;
  
  console.log('Received trial balance request:');
  console.log('- Company ID:', companyId);
  console.log('- Entries count:', trialBalanceEntries?.length);
  console.log('- Sample entry:', trialBalanceEntries?.[0]);
  
  if (!companyId || !trialBalanceEntries || !metadata) {
    res.status(400).json({ 
      success: false, 
      error: 'Missing required fields',
      details: 'companyId, trialBalanceEntries, and metadata are required' 
    });
    return;
  }

  const now = new Date().toISOString();

  // Insert trial balance set
  const setQuery = `
    INSERT INTO trial_balance_sets (
      company_id, file_name, file_hash, upload_date, 
      period_start, period_end, currency, notes
    ) VALUES (?, ?, ?, ?, ?, ?, 'THB', ?)
  `;

  const fileHash = `${metadata.fileName}_${Date.now()}`;
  const periodStart = `${metadata.reportingYear}-01-01`;
  const periodEnd = `${metadata.reportingYear}-12-31`;
  const notes = JSON.stringify(metadata);

  db.run(setQuery, [
    companyId, metadata.fileName, fileHash, now,
    periodStart, periodEnd, notes
  ], function(err) {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to create trial balance set',
        details: err.message 
      });
      return;
    }

    const setId = this.lastID; // Get the auto-generated ID

    // Insert trial balance entries
    const entryQuery = `
      INSERT INTO trial_balance_entries (
        trial_balance_set_id, account_code, account_name, 
        debit_amount, credit_amount, balance, account_type, created_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `;

    let completed = 0;
    const total = trialBalanceEntries.length;

    if (total === 0) {
      res.json({
        success: true,
        trialBalanceSetId: setId,
        message: 'Trial balance set created (no entries)'
      });
      return;
    }

    trialBalanceEntries.forEach(entry => {
      const accountType = getAccountType(entry.accountCode);
      
      // Ensure numeric values are properly converted
      const debitAmount = parseFloat(entry.debitAmount) || 0;
      const creditAmount = parseFloat(entry.creditAmount) || 0;
      const balance = parseFloat(entry.balance) || 0;
      
      db.run(entryQuery, [
        setId,
        entry.accountCode,
        entry.accountName,
        debitAmount,
        creditAmount,
        balance,
        accountType,
        now
      ], (err) => {
        if (err) {
          console.error('Entry insertion error:', err);
          return;
        }

        completed++;
        if (completed === total) {
          res.json({
            success: true,
            trialBalanceSetId: setId,
            entriesCount: total,
            message: 'Trial balance saved successfully'
          });
        }
      });
    });
  });
});

// Get trial balance sets for a company
app.get('/api/trial-balance/company/:companyId', (req, res) => {
  const { companyId } = req.params;

  const query = `
    SELECT tbs.*, COUNT(tbe.id) as entry_count
    FROM trial_balance_sets tbs
    LEFT JOIN trial_balance_entries tbe ON tbs.id = tbe.trial_balance_set_id
    WHERE tbs.company_id = ?
    GROUP BY tbs.id
    ORDER BY tbs.upload_date DESC
  `;

  db.all(query, [companyId], (err, rows) => {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Database error',
        details: err.message 
      });
      return;
    }

    const sets = rows.map(row => ({
      id: row.id,
      companyId: row.company_id,
      fileName: row.file_name,
      uploadDate: new Date(row.upload_date),
      periodStart: row.period_start,
      periodEnd: row.period_end,
      entryCount: row.entry_count,
      metadata: row.notes ? JSON.parse(row.notes) : null
    }));

    res.json({ 
      success: true, 
      data: sets,
      count: sets.length
    });
  });
});

// Get trial balance entries for a set
app.get('/api/trial-balance/:setId/entries', (req, res) => {
  const { setId } = req.params;

  const query = `
    SELECT * FROM trial_balance_entries 
    WHERE trial_balance_set_id = ? 
    ORDER BY account_code
  `;

  db.all(query, [setId], (err, rows) => {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Database error',
        details: err.message 
      });
      return;
    }

    res.json({ 
      success: true, 
      data: rows,
      count: rows.length
    });
  });
});

// Helper function to determine account type
function getAccountType(accountCode) {
  if (!accountCode) return 'unknown';
  
  const code = accountCode.toString();
  if (code.startsWith('1')) return 'asset';
  if (code.startsWith('2')) return 'liability';
  if (code.startsWith('3')) return 'equity';
  if (code.startsWith('4')) return 'revenue';
  if (code.startsWith('5')) return 'expense';
  return 'unknown';
}

// Save generated financial statements
app.post('/api/statements', (req, res) => {
  const { trialBalanceSetId, statements, excelBuffer } = req.body;
  
  console.log('Received request to save statements:', { trialBalanceSetId });
  
  if (!trialBalanceSetId || !statements) {
    res.status(400).json({ 
      success: false, 
      error: 'Missing required fields',
      details: 'trialBalanceSetId and statements are required' 
    });
    return;
  }

  // For now, just acknowledge the request
  // TODO: Implement actual statement storage
  res.json({ 
    success: true, 
    message: 'Statements saved successfully',
    statementId: trialBalanceSetId + '_statements'
  });
});

// File upload route placeholder
app.post('/api/upload', (req, res) => {
  res.json({ 
    success: true, 
    message: 'File upload endpoint ready' 
  });
});

// Error handling
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ 
    success: false, 
    message: 'Something went wrong!' 
  });
});

// 404 handler - fixed the problematic route
app.use((req, res) => {
  res.status(404).json({ 
    success: false, 
    message: 'API route not found' 
  });
});

app.listen(PORT, () => {
  console.log(`🚀 Backend server running on http://localhost:${PORT}`);
  console.log(`📊 Database: SQLite (./financial_statements.db)`);
  console.log(`🔗 CORS enabled for: ${process.env.FRONTEND_URL || 'http://localhost:5173'}`);
  console.log(`📝 API Health: http://localhost:${PORT}/api/health`);
});
