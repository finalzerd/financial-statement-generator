/**
 * Test script for Thai Government DBD (Department of Business Development) Open API
 * Tests the juristic person lookup endpoint with a sample registration number
 * 
 * API Documentation:
 * URL: https://openapi.dbd.go.th/api/v1/juristic_person/{OrganizationJuristicID}
 * Method: GET
 * Returns 8 data items:
 *   1. เลขทะเบียนนิติบุคคล (Registration Number)
 *   2. สถานะนิติบุคคล (Status)
 *   3. ประเภทนิติบุคคล (Type)
 *   4. ทุนจดทะเบียน (Registered Capital)
 *   5. ชื่อนิติบุคคล (Name - Thai/English)
 *   6. ที่ตั้งสำนักงานใหญ่ (Headquarters Address)
 *   7. วันที่จดทะเบียนจัดตั้ง (Registration Date)
 *   8. วัตถุประสงค์ (Objectives with codes)
 */

const https = require('https');

// DBD Open API Configuration
const DBD_API_BASE_URL = 'https://openapi.dbd.go.th/api/v1/juristic_person';
const API_TOKEN = '7anDtBwbBBXUXXViCKrTo3568tTpNtG9'; // API token from DGA

/**
 * Test with a sample registration number
 * Format: 13-digit number
 * Example from documentation: 0105500002383
 */
const TEST_REGISTRATION_NUMBER = '0105500002383'; // From DBD documentation

/**
 * Make GET request to DBD Open API
 * Try different authentication methods to find the correct one
 */
function testDBDAPI(registrationNumber, authMethod = 'all') {
  console.log('=== Testing DBD Open API ===');
  console.log(`Registration Number: ${registrationNumber}`);
  console.log(`Authentication Method: ${authMethod}`);
  
  // Build API URL with registration number as path parameter
  let apiUrl = `${DBD_API_BASE_URL}/${registrationNumber}`;
  
  // Try with token as query parameter
  if (authMethod === 'query' || authMethod === 'all') {
    apiUrl = `${apiUrl}?apikey=${API_TOKEN}`;
  }

  console.log(`API Endpoint: ${apiUrl}`);
  console.log('');

  const url = new URL(apiUrl);
  const headers = {
    'Accept': 'application/json',
    'User-Agent': 'FinancialStatementGenerator/1.0'
  };

  // Add different authentication header formats based on method
  if (authMethod === 'bearer') {
    headers['Authorization'] = `Bearer ${API_TOKEN}`;
  } else if (authMethod === 'apikey') {
    headers['apikey'] = API_TOKEN;
  } else if (authMethod === 'x-api-key') {
    headers['X-API-Key'] = API_TOKEN;
  } else if (authMethod === 'all') {
    // Try all methods at once
    headers['Authorization'] = `Bearer ${API_TOKEN}`;
    headers['X-API-Key'] = API_TOKEN;
    headers['apikey'] = API_TOKEN;
  }

  const options = {
    hostname: url.hostname,
    path: url.pathname + url.search,
    method: 'GET',
    headers: headers
  };

  console.log('Request Headers:', options.headers);
  console.log('Request Method:', options.method);
  console.log('');
  console.log('Sending GET request...\n');

  const req = https.request(options, (res) => {
    console.log(`Status Code: ${res.statusCode}`);
    console.log('Response Headers:', JSON.stringify(res.headers, null, 2));
    console.log('');

    let data = '';

    res.on('data', (chunk) => {
      data += chunk;
    });

    res.on('end', () => {
      console.log('=== RAW RESPONSE ===');
      console.log(data);
      console.log('');

      try {
        const jsonResponse = JSON.parse(data);
        console.log('=== PARSED JSON RESPONSE ===');
        console.log(JSON.stringify(jsonResponse, null, 2));
        console.log('');

        // Extract key information from response
        // Check various possible response structures
        if (jsonResponse.error) {
          console.log('=== API RETURNED ERROR ===');
          console.log('Error:', jsonResponse.error);
          console.log('Message:', jsonResponse.message);
        } else if (jsonResponse.Message) {
          console.log('=== API RETURNED ERROR ===');
          console.log('Message:', jsonResponse.Message);
        } else if (jsonResponse.status && jsonResponse.status.code === '1000') {
          // Success - extract company data from DBD Open API response
          console.log('=== SUCCESS: Company Information Retrieved ===');
          
          // Extract from nested structure: data[0]["cd:OrganizationJuristicPerson"]
          const orgData = jsonResponse.data && jsonResponse.data[0] 
            ? jsonResponse.data[0]['cd:OrganizationJuristicPerson'] 
            : null;
          
          if (orgData) {
            console.log('\n--- Registration Information ---');
            console.log('Registration Number:', orgData['cd:OrganizationJuristicID'] || 'N/A');
            console.log('Status:', orgData['cd:OrganizationJuristicStatus'] || 'N/A');
            console.log('Company Type:', orgData['cd:OrganizationJuristicType'] || 'N/A');
            console.log('Branch:', orgData['cd:OrganizationJuristicBranchName'] || 'N/A');
            
            console.log('\n--- Company Names ---');
            console.log('Thai Name:', orgData['cd:OrganizationJuristicNameTH'] || 'N/A');
            console.log('English Name:', orgData['cd:OrganizationJuristicNameEN'] || 'N/A');
            
            console.log('\n--- Financial Information ---');
            const capital = orgData['cd:OrganizationJuristicRegisterCapital'];
            console.log('Registered Capital:', capital ? `${parseFloat(capital).toLocaleString()} บาท` : 'N/A');
            
            console.log('\n--- Dates ---');
            const regDate = orgData['cd:OrganizationJuristicRegisterDate'];
            if (regDate && regDate.length === 8) {
              // Convert YYYYMMDD to readable format
              const year = regDate.substring(0, 4);
              const month = regDate.substring(4, 6);
              const day = regDate.substring(6, 8);
              console.log('Registration Date:', `${day}/${month}/${year}`);
            } else {
              console.log('Registration Date:', regDate || 'N/A');
            }
            
            console.log('\n--- Location ---');
            const address = orgData['cd:OrganizationJuristicAddress'];
            if (address && address['cr:AddressType']) {
              const addr = address['cr:AddressType'];
              const fullAddress = [
                addr['cd:Address'],
                addr['cd:CitySubDivision']?.['cr:CitySubDivisionTextTH'],
                addr['cd:City']?.['cr:CityTextTH'],
                addr['cd:CountrySubDivision']?.['cr:CountrySubDivisionTextTH']
              ].filter(Boolean).join(', ');
              console.log('Address:', fullAddress || 'N/A');
            } else {
              console.log('Address: N/A');
            }
            
            console.log('\n--- Business Objectives ---');
            const objective = orgData['cd:OrganizationJuristicObjective'];
            if (objective && objective['td:JuristicObjective']) {
              const obj = objective['td:JuristicObjective'];
              console.log('Code:', obj['td:JuristicObjectiveCode'] || 'N/A');
              console.log('Description (TH):', obj['td:JuristicObjectiveTextTH'] || 'N/A');
              console.log('Description (EN):', obj['td:JuristicObjectiveTextEN'] || 'N/A');
            } else {
              console.log('N/A');
            }
          } else {
            console.log('⚠️  No organization data found in response');
          }
        } else {
          console.log('=== UNEXPECTED RESPONSE FORMAT ===');
          console.log('Could not determine success/failure status');
        }
      } catch (parseError) {
        console.error('=== ERROR PARSING JSON ===');
        console.error(parseError.message);
      }
    });
  });

  req.on('error', (error) => {
    console.error('=== REQUEST ERROR ===');
    console.error('Error Type:', error.code);
    console.error('Error Message:', error.message);
    console.error('');
    console.error('Possible causes:');
    console.error('1. Network connectivity issues');
    console.error('2. API endpoint is down or changed');
    console.error('3. Firewall blocking HTTPS requests');
  // No request body for GET request/TLS certificate issues');
  });

  req.on('timeout', () => {
    console.error('=== REQUEST TIMEOUT ===');
    req.abort();
  });

  // Set timeout to 30 seconds
  req.setTimeout(30000);

  // No request body for GET request
  req.end();
}

// Run the test
console.log('╔════════════════════════════════════════════════════════════╗');
console.log('║  DBD API Test Script - Thai Juristic Person Lookup        ║');
console.log('╚════════════════════════════════════════════════════════════╝');
console.log('');

// Allow passing registration number as command line argument
const registrationNumber = process.argv[2] || TEST_REGISTRATION_NUMBER;

if (!registrationNumber || registrationNumber.length !== 13) {
  console.error('⚠️  WARNING: Registration number should be 13 digits');
  console.error(`   Current value: ${registrationNumber} (${registrationNumber.length} digits)`);
  console.error('');
  console.error('Usage: node test-dbd-api.js [13-digit-registration-number]');
  console.error('Example: node test-dbd-api.js 0105564000123');
  console.error('');
  console.error('Proceeding with test anyway...\n');
}

testDBDAPI(registrationNumber);
