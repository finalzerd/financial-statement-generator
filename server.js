import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import sqlite3 from 'sqlite3';
import multer from 'multer';
// Uncomment next line to use PostgreSQL instead of SQLite
// import { DatabaseConfig } from './src/server/config/database.js';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3001;

// Database connection
const dbPath = './financial_statements.db';
const db = new sqlite3.Database(dbPath);

const VALID_DETAIL_ONE_MODES = new Set(['auto', 'service', 'inventory', 'both']);
const DETAIL_SETTINGS_TABLE_SQL = `
  CREATE TABLE IF NOT EXISTS company_detail_settings (
    company_id INTEGER PRIMARY KEY,
    detail_one_mode TEXT DEFAULT 'auto',
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE
  )
`;

// Add share columns to companies table if they don't exist
db.run('ALTER TABLE companies ADD COLUMN number_of_shares INTEGER', (err) => {
  // Ignore error if column already exists
});

db.run('ALTER TABLE companies ADD COLUMN share_value REAL', (err) => {
  // Ignore error if column already exists
});

// Ensure company_account_mappings table exists
db.run(`
  CREATE TABLE IF NOT EXISTS company_account_mappings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id INTEGER NOT NULL,
    note_type TEXT NOT NULL,
    note_number INTEGER,
    note_title TEXT,
    account_ranges TEXT NOT NULL,
    sub_category_rules TEXT, -- JSON blob for optional sub-category rules (currently only used for cash)
    is_active INTEGER DEFAULT 1,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(company_id, note_type),
    FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE CASCADE
  )
`, (err) => {
  if (err) {
    console.error('Error ensuring company_account_mappings table exists:', err);
  }
});

// Add sub_category_rules column if it doesn't exist (for existing databases)
db.run('ALTER TABLE company_account_mappings ADD COLUMN sub_category_rules TEXT', (err) => {
  // Ignore error if column already exists
});

// Ensure company_detail_settings table exists for detail note preferences
db.run(DETAIL_SETTINGS_TABLE_SQL, (err) => {
  if (err) {
    console.error('Error ensuring company_detail_settings table exists:', err);
  }
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
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

// Multer setup for multipart/form-data (used when saving generated statements with Excel file)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 } // 5MB limit for uploaded Excel file
});

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

// ============== DETAIL SETTINGS ENDPOINTS ==============

app.get('/api/companies/:companyId/detail-settings', (req, res) => {
  const companyId = parseInt(req.params.companyId, 10);
  const defaultMode = 'auto';

  if (Number.isNaN(companyId)) {
    res.status(400).json({
      success: false,
      error: 'Invalid company ID'
    });
    return;
  }

  db.get('SELECT id FROM companies WHERE id = ?', [companyId], (err, company) => {
    if (err) {
      console.error('Database error when fetching company for detail settings:', err);
      res.status(500).json({
        success: false,
        error: 'Database error',
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

    const sendDefault = () => {
      const now = new Date().toISOString();
      const insertQuery = `
        INSERT INTO company_detail_settings (company_id, detail_one_mode, updated_at)
        VALUES (?, ?, ?)
      `;

      db.run(insertQuery, [companyId, defaultMode, now], (insertErr) => {
        if (insertErr && !(insertErr.message && insertErr.message.includes('UNIQUE constraint failed'))) {
          console.error('Database error when initializing detail settings:', insertErr);
          res.status(500).json({
            success: false,
            error: 'Failed to initialize detail settings',
            details: insertErr.message
          });
          return;
        }

        res.json({
          detailOneMode: defaultMode,
          updatedAt: now
        });
      });
    };

    db.get(
      'SELECT detail_one_mode, updated_at FROM company_detail_settings WHERE company_id = ?',
      [companyId],
      (settingsErr, row) => {
        if (settingsErr) {
          if (settingsErr.message && settingsErr.message.includes('no such table')) {
            db.run(DETAIL_SETTINGS_TABLE_SQL, (createErr) => {
              if (createErr) {
                console.error('Failed to create detail settings table on-demand:', createErr);
                res.status(500).json({
                  success: false,
                  error: 'Database error',
                  details: createErr.message
                });
                return;
              }
              sendDefault();
            });
            return;
          }

          console.error('Database error when fetching detail settings:', settingsErr);
          res.status(500).json({
            success: false,
            error: 'Database error',
            details: settingsErr.message
          });
          return;
        }

        if (row) {
          res.json({
            detailOneMode: row.detail_one_mode || defaultMode,
            updatedAt: row.updated_at
          });
          return;
        }

        sendDefault();
      }
    );
  });
});

app.put('/api/companies/:companyId/detail-settings', (req, res) => {
  const companyId = parseInt(req.params.companyId, 10);
  const { detailOneMode } = req.body;

  if (Number.isNaN(companyId)) {
    res.status(400).json({
      success: false,
      error: 'Invalid company ID'
    });
    return;
  }

  if (typeof detailOneMode !== 'string' || !VALID_DETAIL_ONE_MODES.has(detailOneMode)) {
    res.status(400).json({
      success: false,
      error: 'Invalid detailOneMode value',
      details: `Allowed values: ${Array.from(VALID_DETAIL_ONE_MODES).join(', ')}`
    });
    return;
  }

  db.get('SELECT id FROM companies WHERE id = ?', [companyId], (err, company) => {
    if (err) {
      console.error('Database error when validating company for detail settings update:', err);
      res.status(500).json({
        success: false,
        error: 'Database error',
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

    const timestamp = new Date().toISOString();
    const upsertQuery = `
      INSERT INTO company_detail_settings (company_id, detail_one_mode, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(company_id) DO UPDATE SET
        detail_one_mode = excluded.detail_one_mode,
        updated_at = excluded.updated_at
    `;

    db.run(upsertQuery, [companyId, detailOneMode, timestamp], function(updateErr) {
      if (updateErr) {
        console.error('Database error when updating detail settings:', updateErr);
        res.status(500).json({
          success: false,
          error: 'Failed to update detail settings',
          details: updateErr.message
        });
        return;
      }

      res.json({
        success: true,
        detailOneMode,
        updatedAt: timestamp
      });
    });
  });
});

// ============== ACCOUNT MAPPING ENDPOINTS ==============

// Get all account mappings for a company
app.get('/api/companies/:companyId/account-mappings', (req, res) => {
  const { companyId } = req.params;

  const query = `
    SELECT * FROM company_account_mappings 
    WHERE company_id = ? 
    ORDER BY note_number, note_type
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

    const mappings = rows.map(row => ({
      id: row.id,
      companyId: row.company_id,
      noteType: row.note_type,
      noteNumber: row.note_number,
      noteTitle: row.note_title,
      accountRanges: JSON.parse(row.account_ranges),
      subCategoryRules: row.sub_category_rules ? JSON.parse(row.sub_category_rules) : null,
      isActive: Boolean(row.is_active),
      createdAt: row.created_at,
      updatedAt: row.updated_at
    }));

    res.json(mappings);
  });
});

// Update account mapping for a specific note type
app.put('/api/companies/:companyId/account-mappings/:noteType', (req, res) => {
  const { companyId, noteType } = req.params;
  const { noteNumber, noteTitle, accountRanges, isActive, subCategoryRules } = req.body;

  if (!accountRanges) {
    res.status(400).json({ 
      success: false, 
      error: 'Missing required fields',
      details: 'accountRanges is required' 
    });
    return;
  }

  const now = new Date().toISOString();
  const query = `
    UPDATE company_account_mappings 
    SET note_number = ?, note_title = ?, account_ranges = ?, sub_category_rules = ?, is_active = ?, updated_at = ?
    WHERE company_id = ? AND note_type = ?
  `;

  db.run(query, [
    noteNumber,
    noteTitle,
    JSON.stringify(accountRanges),
    subCategoryRules ? JSON.stringify(subCategoryRules) : null,
    isActive ? 1 : 0,
    now,
    companyId,
    noteType
  ], function(err) {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to update account mapping',
        details: err.message 
      });
      return;
    }

    if (this.changes === 0) {
      res.status(404).json({ 
        success: false, 
        error: 'Account mapping not found',
        details: `No mapping found for company ${companyId} and note type ${noteType}` 
      });
      return;
    }

    res.json({ 
      success: true, 
      message: 'Account mapping updated successfully' 
    });
  });
});

// Create new account mapping
app.post('/api/companies/:companyId/account-mappings', (req, res) => {
  const { companyId } = req.params;
  const { noteType, noteNumber, noteTitle, accountRanges, isActive, subCategoryRules } = req.body;

  if (!noteType || !accountRanges) {
    res.status(400).json({ 
      success: false, 
      error: 'Missing required fields',
      details: 'noteType and accountRanges are required' 
    });
    return;
  }

  const now = new Date().toISOString();
  const query = `
    INSERT INTO company_account_mappings 
    (company_id, note_type, note_number, note_title, account_ranges, sub_category_rules, is_active, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `;

  db.run(query, [
    companyId,
    noteType,
    noteNumber,
    noteTitle,
    JSON.stringify(accountRanges),
    subCategoryRules ? JSON.stringify(subCategoryRules) : null,
    isActive ? 1 : 0,
    now,
    now
  ], function(err) {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to create account mapping',
        details: err.message 
      });
      return;
    }

    res.status(201).json({ 
      success: true, 
      id: this.lastID,
      message: 'Account mapping created successfully' 
    });
  });
});

// Delete account mapping
app.delete('/api/companies/:companyId/account-mappings/:noteType', (req, res) => {
  const { companyId, noteType } = req.params;

  const query = `DELETE FROM company_account_mappings WHERE company_id = ? AND note_type = ?`;

  db.run(query, [companyId, noteType], function(err) {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to delete account mapping',
        details: err.message 
      });
      return;
    }

    if (this.changes === 0) {
      res.status(404).json({ 
        success: false, 
        error: 'Account mapping not found',
        details: `No mapping found for company ${companyId} and note type ${noteType}` 
      });
      return;
    }

    res.json({ 
      success: true, 
      message: 'Account mapping deleted successfully' 
    });
  });
});

// Reset account mappings to default
app.post('/api/companies/:companyId/account-mappings/reset', (req, res) => {
  const { companyId } = req.params;

  // Import the default mappings from our migration script
  const DEFAULT_MAPPINGS = [
    {
      noteType: 'cash',
      noteNumber: 7,
      noteTitle: 'เงินสดและรายการเทียบเท่าเงินสด',
      accountRanges: JSON.stringify({
        ranges: [{ from: 1000, to: 1099 }]
      })
    },
    {
      noteType: 'receivables',
      noteNumber: 8,
      noteTitle: 'ลูกหนี้การค้าและลูกหนี้อื่น',
      accountRanges: JSON.stringify({
        ranges: [{ from: 1140, to: 1215 }]
      })
    },
    {
      noteType: 'inventory',
      noteNumber: 9,
      noteTitle: 'สินค้าคงเหลือ',
      accountRanges: JSON.stringify({
        includes: [1510]
      })
    },
    {
      noteType: 'prepaid_expenses',
      noteNumber: 10,
      noteTitle: 'ค่าใช้จ่ายจ่ายล่วงหน้า',
      accountRanges: JSON.stringify({
        ranges: [{ from: 1400, to: 1439 }]
      })
    },
    {
      noteType: 'ppe_cost',
      noteNumber: 11,
      noteTitle: 'ที่ดิน อาคาร และอุปกรณ์',
      accountRanges: JSON.stringify({
        ranges: [{ from: 1610, to: 1659 }]
      })
    },
    {
      noteType: 'other_assets',
      noteNumber: 12,
      noteTitle: 'สินทรัพย์อื่น',
      accountRanges: JSON.stringify({
        ranges: [{ from: 1660, to: 1700 }]
      })
    },
    {
      noteType: 'asset_short_term_loans',
      noteNumber: 9,
      noteTitle: 'เงินให้กู้ยืมระยะสั้น',
      accountRanges: JSON.stringify({
        includes: [1141]
      })
    },
    {
      noteType: 'asset_long_term_loans',
      noteNumber: 13,
      noteTitle: 'เงินให้กู้ยืมระยะยาว',
      accountRanges: JSON.stringify({
        includes: [1710]
      })
    },
    {
      noteType: 'other_income',
      noteNumber: 14,
      noteTitle: 'รายได้อื่น',
      accountRanges: JSON.stringify({
        ranges: [{ from: 4110, to: 4999 }]
      })
    },
    {
      noteType: 'detail_service_costs',
      noteNumber: 1,
      noteTitle: 'ต้นทุนการให้บริการ (รายละเอียดที่ 1)',
      accountRanges: JSON.stringify({
        ranges: [{ from: 5000, to: 5099 }]
      })
    },
    {
      noteType: 'bank_overdrafts',
      noteNumber: 15,
      noteTitle: 'เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน',
      accountRanges: JSON.stringify({
        ranges: [{ from: 2001, to: 2009 }]
      })
    },
    {
      noteType: 'payables',
      noteNumber: 16,
      noteTitle: 'เจ้าหนี้การค้าและเจ้าหนี้อื่น',
      accountRanges: JSON.stringify({
        ranges: [{ from: 2010, to: 2999 }],
        excludes: [2030, 2045, 2050, 2051, 2052, 2100, 2101, 2102, 2103, 2120, 2121, 2122, 2123]
      })
    },
    {
      noteType: 'short_term_loans',
      noteNumber: 17,
      noteTitle: 'เงินกู้ยืมระยะสั้น',
      accountRanges: JSON.stringify({
        includes: [2030]
      })
    },
    {
      noteType: 'income_tax_payable',
      noteNumber: 18,
      noteTitle: 'ภาษีเงินได้นิติบุคคลค้างจ่าย',
      accountRanges: JSON.stringify({
        includes: [2045]
      })
    },
    {
      noteType: 'long_term_loans_fi',
      noteNumber: 19,
      noteTitle: 'เงินกู้ยืมระยะยาวจากสถาบันการเงิน',
      accountRanges: JSON.stringify({
        ranges: [{ from: 2120, to: 2123 }],
        excludes: [2121]
      })
    },
    {
      noteType: 'hire_purchase_creditors',
      noteNumber: 21,
      noteTitle: 'เจ้าหนี้ตามสัญญาเช่าซื้อ',
      accountRanges: JSON.stringify({
        includes: [2015, 1644.2, 1644.1]
      })
    },
    {
      noteType: 'long_term_loans_other',
      noteNumber: 20,
      noteTitle: 'เงินกู้ยืมระยะยาวอื่น',
      accountRanges: JSON.stringify({
        includes: [2050, 2051, 2052, 2100, 2101, 2102, 2103]
      })
    }
  ];

  // First, delete existing mappings
  db.run('DELETE FROM company_account_mappings WHERE company_id = ?', [companyId], (err) => {
    if (err) {
      console.error('Database error:', err);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to delete existing mappings',
        details: err.message 
      });
      return;
    }

    // Then, insert default mappings
    const now = new Date().toISOString();
    let completed = 0;
    const total = DEFAULT_MAPPINGS.length;

    if (total === 0) {
      res.json({
        success: true,
        message: 'Account mappings reset (no defaults to insert)'
      });
      return;
    }

    DEFAULT_MAPPINGS.forEach(mapping => {
      const insertQuery = `
        INSERT INTO company_account_mappings 
        (company_id, note_type, note_number, note_title, account_ranges, sub_category_rules, is_active, created_at, updated_at) 
        VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?)
      `;

      db.run(insertQuery, [
        companyId,
        mapping.noteType,
        mapping.noteNumber,
        mapping.noteTitle,
        mapping.accountRanges,
        mapping.noteType === 'cash' ? null : null,
        now,
        now
      ], (err) => {
        if (err) {
          console.error(`Error inserting mapping ${mapping.noteType}:`, err);
          return;
        }

        completed++;
        if (completed === total) {
          const detailResetQuery = `
            INSERT INTO company_detail_settings (company_id, detail_one_mode, updated_at)
            VALUES (?, 'auto', ?)
            ON CONFLICT(company_id) DO UPDATE SET
              detail_one_mode = 'auto',
              updated_at = excluded.updated_at
          `;

          db.run(detailResetQuery, [companyId, now], (detailErr) => {
            if (detailErr) {
              console.error('Error resetting detail settings during account mapping reset:', detailErr);
              res.status(500).json({
                success: false,
                error: 'Failed to reset detail settings',
                details: detailErr.message
              });
              return;
            }

            res.json({
              success: true,
              message: `Account mappings reset to default (${total} mappings created)`
            });
          });
        }
      });
    });
  });
});

// Validate account mappings against trial balance data
app.post('/api/companies/:companyId/validate-mappings', (req, res) => {
  const { companyId } = req.params;
  const { trialBalanceData } = req.body;

  if (!trialBalanceData || !Array.isArray(trialBalanceData)) {
    res.status(400).json({ 
      success: false, 
      error: 'Missing or invalid trial balance data',
      details: 'trialBalanceData array is required' 
    });
    return;
  }

  // Get account mappings for the company
  const query = `
    SELECT * FROM company_account_mappings 
    WHERE company_id = ? AND is_active = 1
    ORDER BY note_number, note_type
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

    // Process validation logic
    const mappings = rows.map(row => ({
      noteType: row.note_type,
      accountRanges: JSON.parse(row.account_ranges)
    }));

    const mappedAccountCodes = new Set();
    const accountToNotes = new Map();
    const warnings = [];
    const errors = [];

    // Check each mapping against trial balance
    mappings.forEach(mapping => {
      const { accountRanges } = mapping;
      
      trialBalanceData.forEach(account => {
        const code = parseFloat(account.accountCode || '0');
        let isMatched = false;

        // Check ranges
        if (accountRanges.ranges) {
          isMatched = accountRanges.ranges.some(range => 
            code >= range.from && code <= range.to
          );
        }

        // Check includes
        if (accountRanges.includes) {
          isMatched = isMatched || accountRanges.includes.includes(code);
        }

        // Check excludes
        if (accountRanges.excludes && isMatched) {
          isMatched = !accountRanges.excludes.includes(code);
        }

        if (isMatched) {
          mappedAccountCodes.add(account.accountCode);
          
          if (!accountToNotes.has(account.accountCode)) {
            accountToNotes.set(account.accountCode, []);
          }
          accountToNotes.get(account.accountCode).push(mapping.noteType);
        }
      });
    });

    // Find unmapped accounts
    const unmappedAccounts = trialBalanceData.filter(account => 
      !mappedAccountCodes.has(account.accountCode)
    ).map(account => ({
      accountCode: account.accountCode,
      accountName: account.accountName,
      balance: account.balance || account.currentBalance || 0
    }));

    // Find conflicting accounts (mapped to multiple notes)
    const conflictingAccounts = [];
    accountToNotes.forEach((noteTypes, accountCode) => {
      if (noteTypes.length > 1) {
        const account = trialBalanceData.find(a => a.accountCode === accountCode);
        if (account) {
          conflictingAccounts.push({
            accountCode,
            accountName: account.accountName,
            noteTypes
          });
        }
      }
    });

    // Add warnings and errors
    if (unmappedAccounts.length > 0) {
      warnings.push(`${unmappedAccounts.length} accounts are not mapped to any note`);
    }

    if (conflictingAccounts.length > 0) {
      errors.push(`${conflictingAccounts.length} accounts are mapped to multiple notes`);
    }

    const totalAccounts = trialBalanceData.length;
    const mappedAccounts = mappedAccountCodes.size;
    const coveragePercentage = totalAccounts > 0 ? (mappedAccounts / totalAccounts) * 100 : 0;

    res.json({
      isValid: errors.length === 0,
      warnings,
      errors,
      unmappedAccounts,
      conflictingAccounts,
      mappingCoverage: {
        totalAccounts,
        mappedAccounts,
        unmappedAccounts: unmappedAccounts.length,
        coveragePercentage
      }
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

// Save generated financial statements (supports multipart/form-data or JSON)
app.post('/api/statements', upload.single('excelFile'), (req, res) => {
  try {
    // Support both multipart/form-data (FormData) and application/json
    let trialBalanceSetId;
    let statementsRaw;

    if (req.is('multipart/form-data')) {
      trialBalanceSetId = req.body?.trialBalanceSetId;
      statementsRaw = req.body?.statements; // JSON string
    } else {
      // JSON request
      trialBalanceSetId = req.body?.trialBalanceSetId;
      statementsRaw = req.body?.statements;
    }

    let parsedStatements = null;
    if (typeof statementsRaw === 'string') {
      try {
        parsedStatements = JSON.parse(statementsRaw);
      } catch (e) {
        console.warn('Failed to parse statements JSON string');
      }
    } else if (statementsRaw && typeof statementsRaw === 'object') {
      parsedStatements = statementsRaw;
    }

    console.log('Received request to save statements:', { trialBalanceSetId, hasFile: !!req.file });

    if (!trialBalanceSetId || !parsedStatements) {
      res.status(400).json({
        success: false,
        error: 'Missing required fields',
        details: 'trialBalanceSetId and statements are required'
      });
      return;
    }

    // For now, just acknowledge the request - future persistence can store req.file.buffer
    res.json({
      success: true,
      message: 'Statements saved successfully',
      statementId: trialBalanceSetId + '_statements'
    });
  } catch (err) {
    console.error('Error saving statements:', err);
    res.status(500).json({
      success: false,
      error: 'Internal server error',
      details: err.message
    });
  }
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
