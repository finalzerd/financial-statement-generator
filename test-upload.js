// Test script to verify upload functionality
import fs from 'fs';
import path from 'path';

async function testUpload() {
  try {
    console.log('🧪 Testing CSV upload functionality...');
    
    // Read the CSV file
    const csvContent = fs.readFileSync(new URL('TB01.csv', import.meta.url), 'utf8');
    console.log('📁 CSV file loaded:', csvContent.split('\n').length, 'lines');
    
    // Test data structure
    const testData = {
      companyId: "2", // Using the company ID from our API test
      trialBalanceEntries: [
        {
          accountCode: "1010",
          accountName: "เงินสด",
          debitAmount: 89932.38,
          creditAmount: 0,
          balance: 89932.38
        },
        {
          accountCode: "4010",
          accountName: "รายได้ค่าบริการ",
          debitAmount: 0,
          creditAmount: 125000,
          balance: -125000
        }
      ],
      metadata: {
        fileName: "TB01.csv",
        processingType: "single-year",
        reportingYear: 2024,
        reportingPeriod: "annual",
        totalEntries: 2
      }
    };
    
    // Send POST request
    const response = await fetch('http://localhost:3001/api/trial-balance', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(testData)
    });
    
    const result = await response.json();
    console.log('📊 Upload result:', result);
    
    if (result.success) {
      console.log('✅ Upload test PASSED!');
      console.log(`📈 Trial Balance Set ID: ${result.trialBalanceSetId}`);
      console.log(`📋 Entries saved: ${result.entriesCount}`);
    } else {
      console.log('❌ Upload test FAILED:', result.error);
    }
    
  } catch (error) {
    console.error('❌ Test error:', error.message);
  }
}

testUpload();
