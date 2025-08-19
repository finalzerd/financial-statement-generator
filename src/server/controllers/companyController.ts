import type { Request, Response } from 'express';
import { DatabaseService } from '../services/databaseService';

export class CompanyController {
  
  /**
   * Create a new company
   */
  static async createCompany(req: Request, res: Response) {
    try {
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

      // Validation
      if (!name || !type) {
        return res.status(400).json({ 
          success: false,
          error: 'Company name and type are required',
          details: 'Both name and type fields must be provided'
        });
      }

      if (!['ห้างหุ้นส่วนจำกัด', 'บริษัทจำกัด'].includes(type)) {
        return res.status(400).json({
          success: false,
          error: 'Invalid company type',
          details: 'Type must be either "ห้างหุ้นส่วนจำกัด" or "บริษัทจำกัด"'
        });
      }

      const company = await DatabaseService.createCompany({
        name: name.trim(),
        type,
        registrationNumber: registrationNumber?.trim(),
        address: address?.trim(),
        businessDescription: businessDescription?.trim(),
        taxId: taxId?.trim(),
        defaultReportingYear: defaultReportingYear || new Date().getFullYear(),
        numberOfShares: type === 'บริษัทจำกัด' ? numberOfShares : undefined,
        shareValue: type === 'บริษัทจำกัด' ? shareValue : undefined
      });

      res.status(201).json({ 
        success: true, 
        company,
        message: `Company "${company.name}" created successfully`
      });

    } catch (error) {
      console.error('❌ Error creating company:', error);
      
      if (error instanceof Error && error.message.includes('duplicate')) {
        return res.status(409).json({
          success: false,
          error: 'Company already exists',
          details: 'A company with this information already exists'
        });
      }

      res.status(500).json({ 
        success: false,
        error: 'Failed to create company',
        details: error instanceof Error ? error.message : 'Unknown server error'
      });
    }
  }

  /**
   * Get all companies
   */
  static async getCompanies(_req: Request, res: Response) {
    try {
      const companies = await DatabaseService.getCompanies();
      
      res.json({ 
        success: true, 
        companies,
        count: companies.length
      });

    } catch (error) {
      console.error('❌ Error fetching companies:', error);
      res.status(500).json({ 
        success: false,
        error: 'Failed to fetch companies',
        details: error instanceof Error ? error.message : 'Unknown server error'
      });
    }
  }

  /**
   * Get company by ID
   */
  static async getCompanyById(req: Request, res: Response) {
    try {
      const { id } = req.params;

      if (!id) {
        return res.status(400).json({
          success: false,
          error: 'Company ID is required',
          details: 'Please provide a valid company ID'
        });
      }

      const company = await DatabaseService.getCompanyById(id);

      if (!company) {
        return res.status(404).json({ 
          success: false,
          error: 'Company not found',
          details: `No company found with ID: ${id}`
        });
      }

      res.json({ 
        success: true, 
        company 
      });

    } catch (error) {
      console.error('❌ Error fetching company:', error);
      res.status(500).json({ 
        success: false,
        error: 'Failed to fetch company',
        details: error instanceof Error ? error.message : 'Unknown server error'
      });
    }
  }

  /**
   * Update company
   */
  static async updateCompany(req: Request, res: Response) {
    try {
      const { id } = req.params;
      const updates = req.body;

      if (!id) {
        return res.status(400).json({
          success: false,
          error: 'Company ID is required',
          details: 'Please provide a valid company ID'
        });
      }

      // Clean up the updates object
      const cleanUpdates = Object.fromEntries(
        Object.entries(updates).filter(([_, value]) => value !== undefined && value !== null)
      );

      if (Object.keys(cleanUpdates).length === 0) {
        return res.status(400).json({
          success: false,
          error: 'No valid updates provided',
          details: 'Please provide at least one field to update'
        });
      }

      const company = await DatabaseService.updateCompany(id, cleanUpdates);

      if (!company) {
        return res.status(404).json({
          success: false,
          error: 'Company not found',
          details: `No company found with ID: ${id}`
        });
      }

      res.json({
        success: true,
        company,
        message: 'Company updated successfully'
      });

    } catch (error) {
      console.error('❌ Error updating company:', error);
      res.status(500).json({
        success: false,
        error: 'Failed to update company',
        details: error instanceof Error ? error.message : 'Unknown server error'
      });
    }
  }

  /**
   * Get company's trial balance history
   */
  static async getCompanyHistory(req: Request, res: Response) {
    try {
      const { id } = req.params;

      if (!id) {
        return res.status(400).json({
          success: false,
          error: 'Company ID is required',
          details: 'Please provide a valid company ID'
        });
      }

      // First, verify company exists
      const company = await DatabaseService.getCompanyById(id);
      if (!company) {
        return res.status(404).json({
          success: false,
          error: 'Company not found',
          details: `No company found with ID: ${id}`
        });
      }

      const trialBalanceSets = await DatabaseService.getTrialBalanceHistory(id);

      res.json({ 
        success: true, 
        company: {
          id: company.id,
          name: company.name,
          type: company.type
        },
        trialBalanceSets,
        count: trialBalanceSets.length
      });

    } catch (error) {
      console.error('❌ Error fetching company history:', error);
      res.status(500).json({ 
        success: false,
        error: 'Failed to fetch company history',
        details: error instanceof Error ? error.message : 'Unknown server error'
      });
    }
  }

  /**
   * Get company statistics
   */
  static async getCompanyStats(req: Request, res: Response) {
    try {
      const { id } = req.params;

      if (!id) {
        return res.status(400).json({
          success: false,
          error: 'Company ID is required'
        });
      }

      const company = await DatabaseService.getCompanyById(id);
      if (!company) {
        return res.status(404).json({
          success: false,
          error: 'Company not found'
        });
      }

      const trialBalanceSets = await DatabaseService.getTrialBalanceHistory(id);
      
      const stats = {
        totalTrialBalances: trialBalanceSets.length,
        yearsCovered: [...new Set(trialBalanceSets.map(tb => tb.reportingYear))].sort(),
        latestUpload: trialBalanceSets.length > 0 ? trialBalanceSets[0].uploadedAt : null,
        processingTypes: [...new Set(trialBalanceSets.map(tb => tb.processingType))]
      };

      res.json({
        success: true,
        company: {
          id: company.id,
          name: company.name,
          type: company.type
        },
        stats
      });

    } catch (error) {
      console.error('❌ Error fetching company stats:', error);
      res.status(500).json({
        success: false,
        error: 'Failed to fetch company statistics',
        details: error instanceof Error ? error.message : 'Unknown server error'
      });
    }
  }
}
