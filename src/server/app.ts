import express from 'express';
import cors from 'cors';
import fileUpload from 'express-fileupload';
import dotenv from 'dotenv';
import { DatabaseConfig } from './config/database';
import companyRoutes from './routes/companies';

// Load environment variables
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3001;

// ============== MIDDLEWARE ==============

// CORS configuration
app.use(cors({
  origin: process.env.FRONTEND_URL || 'http://localhost:5173',
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

// Body parsing middleware
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// File upload middleware
app.use(fileUpload({
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB limit
  abortOnLimit: true,
  responseOnLimit: 'File size limit exceeded (50MB max)',
  parseNested: true
}));

// Request logging middleware
app.use((req, _res, next) => {
  const timestamp = new Date().toISOString();
  console.log(`[${timestamp}] ${req.method} ${req.path}`);
  next();
});

// ============== ROUTES ==============

// Health check endpoint
app.get('/api/health', (_req, res) => {
  res.json({ 
    status: 'OK', 
    timestamp: new Date().toISOString(),
    service: 'Financial Statement Generator API',
    version: '1.0.0'
  });
});

// API routes
app.use('/api/companies', companyRoutes);

// 404 handler for unknown routes
app.use('*', (_req, res) => {
  res.status(404).json({
    success: false,
    error: 'Route not found',
    details: 'The requested endpoint does not exist'
  });
});

// ============== ERROR HANDLING ==============

// Global error handler
app.use((error: any, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  console.error('❌ Global error handler:', error);
  
  // Handle specific error types
  if (error.code === 'LIMIT_FILE_SIZE') {
    return res.status(413).json({
      success: false,
      error: 'File too large',
      details: 'Maximum file size is 50MB'
    });
  }

  if (error.type === 'entity.parse.failed') {
    return res.status(400).json({
      success: false,
      error: 'Invalid JSON',
      details: 'Request body contains invalid JSON'
    });
  }

  // Generic error response
  res.status(500).json({
    success: false,
    error: 'Internal server error',
    details: process.env.NODE_ENV === 'development' ? error.message : 'An unexpected error occurred'
  });
});

// ============== SERVER STARTUP ==============

async function startServer() {
  try {
    console.log('🚀 Starting Financial Statement Generator API...');
    console.log(`📍 Environment: ${process.env.NODE_ENV || 'development'}`);
    
    // Initialize database connection
    console.log('🔗 Initializing database connection...');
    DatabaseConfig.initialize();
    
    // Test database connection
    const dbConnected = await DatabaseConfig.testConnection();
    if (!dbConnected) {
      throw new Error('Failed to connect to database');
    }
    
    // Create tables if they don't exist
    console.log('🏗️  Setting up database tables...');
    await DatabaseConfig.createTables();
    
    // Start the server
    app.listen(PORT, () => {
      console.log('');
      console.log('✅ =================================');
      console.log('✅ Server started successfully!');
      console.log(`✅ API Server: http://localhost:${PORT}`);
      console.log(`✅ Frontend: ${process.env.FRONTEND_URL || 'http://localhost:5173'}`);
      console.log(`✅ Health check: http://localhost:${PORT}/api/health`);
      console.log('✅ =================================');
      console.log('');
      console.log('📊 Available endpoints:');
      console.log('   POST   /api/companies                 - Create company');
      console.log('   GET    /api/companies                 - List companies');
      console.log('   GET    /api/companies/:id             - Get company');
      console.log('   PUT    /api/companies/:id             - Update company');
      console.log('   GET    /api/companies/:id/trial-balances - Company history');
      console.log('   GET    /api/companies/:id/stats       - Company statistics');
      console.log('');
    });

  } catch (error) {
    console.error('❌ Failed to start server:', error);
    console.error('💡 Please check your database configuration in .env file');
    process.exit(1);
  }
}

// Handle graceful shutdown
process.on('SIGINT', async () => {
  console.log('\n🛑 Shutting down server...');
  try {
    await DatabaseConfig.close();
    console.log('✅ Database connection closed');
    process.exit(0);
  } catch (error) {
    console.error('❌ Error during shutdown:', error);
    process.exit(1);
  }
});

process.on('SIGTERM', async () => {
  console.log('\n🛑 Received SIGTERM, shutting down gracefully...');
  try {
    await DatabaseConfig.close();
    process.exit(0);
  } catch (error) {
    console.error('❌ Error during shutdown:', error);
    process.exit(1);
  }
});

// Start the server
startServer();

export default app;
