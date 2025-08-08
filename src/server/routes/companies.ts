import { Router } from 'express';
import { CompanyController } from '../controllers/companyController';

const router = Router();

// Company CRUD operations
router.post('/', CompanyController.createCompany);           // POST /api/companies
router.get('/', CompanyController.getCompanies);             // GET /api/companies
router.get('/:id', CompanyController.getCompanyById);        // GET /api/companies/:id
router.put('/:id', CompanyController.updateCompany);         // PUT /api/companies/:id

// Company-specific data
router.get('/:id/trial-balances', CompanyController.getCompanyHistory);  // GET /api/companies/:id/trial-balances
router.get('/:id/stats', CompanyController.getCompanyStats);             // GET /api/companies/:id/stats

export default router;
