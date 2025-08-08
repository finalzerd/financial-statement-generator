// Database setup script for Financial Statement Generator
import dotenv from 'dotenv';
import { DatabaseConfig } from '../src/server/config/database.js';

// Load environment variables
dotenv.config();

async function setupDatabase() {
  console.log('🚀 Setting up Financial Statement Generator Database...');
  console.log('');
  
  try {
    // Initialize database connection
    console.log('🔗 Connecting to database...');
    DatabaseConfig.initialize();
    
    // Test connection
    const connected = await DatabaseConfig.testConnection();
    if (!connected) {
      throw new Error('Failed to connect to database');
    }
    
    // Create tables
    console.log('🏗️  Creating database tables...');
    await DatabaseConfig.createTables();
    
    console.log('');
    console.log('✅ Database setup completed successfully!');
    console.log('');
    console.log('📋 Created tables:');
    console.log('   • companies');
    console.log('   • trial_balance_sets');
    console.log('   • trial_balance_entries');
    console.log('   • generated_statements');
    console.log('   • company_settings');
    console.log('');
    console.log('🎯 Your Financial Statement Generator is ready to use!');
    console.log('');
    console.log('Next steps:');
    console.log('1. Start the development server: npm run dev');
    console.log('2. Open http://localhost:5173 in your browser');
    console.log('3. Begin creating companies and uploading trial balances');
    
  } catch (error) {
    console.error('❌ Database setup failed:', error);
    console.error('');
    console.error('💡 Troubleshooting tips:');
    console.error('1. Check your .env file configuration');
    console.error('2. Ensure PostgreSQL is running');
    console.error('3. Verify database credentials');
    console.error('4. Make sure the database exists');
    process.exit(1);
  } finally {
    await DatabaseConfig.close();
    process.exit(0);
  }
}

// Run the setup
setupDatabase().catch(console.error);
