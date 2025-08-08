import { SQLiteConfig } from '../src/server/config/sqlite-database';

async function setupDatabase() {
  try {
    console.log('🚀 Starting database setup...');
    
    // Initialize SQLite database
    const db = SQLiteConfig.initialize();
    
    // Test connection
    const isConnected = await SQLiteConfig.testConnection();
    if (!isConnected) {
      throw new Error('Database connection failed');
    }
    
    // Create tables
    await SQLiteConfig.createTables();
    
    console.log('✅ Database setup completed successfully!');
    console.log('📊 Tables created:');
    console.log('   - companies');
    console.log('   - trial_balance_sets'); 
    console.log('   - trial_balance_entries');
    console.log('   - generated_statements');
    console.log('   - company_settings');
    
    // Close connection
    await SQLiteConfig.close();
    
  } catch (error) {
    console.error('❌ Database setup failed:', error);
    process.exit(1);
  }
}

setupDatabase();
