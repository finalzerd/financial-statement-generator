import { DatabaseConfig } from '../config/database';

export type DetailOneMode = 'auto' | 'service' | 'inventory' | 'both';

export interface DetailSettingsRecord {
  detailOneMode: DetailOneMode;
  updatedAt: Date;
  persisted: boolean;
}

const VALID_DETAIL_ONE_MODES: DetailOneMode[] = ['auto', 'service', 'inventory', 'both'];
const VALID_MODE_SET = new Set<DetailOneMode>(VALID_DETAIL_ONE_MODES);

export class DetailSettingsService {
  private static tableInitialized = false;

  private static get pool() {
    return DatabaseConfig.getPool();
  }

  private static async ensureTableExists(): Promise<void> {
    if (this.tableInitialized) {
      return;
    }

    await this.pool.query(`
      CREATE TABLE IF NOT EXISTS company_detail_settings (
        company_id UUID PRIMARY KEY REFERENCES companies(id) ON DELETE CASCADE,
        detail_one_mode VARCHAR(20) NOT NULL DEFAULT 'auto',
        updated_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
      )
    `);

    this.tableInitialized = true;
  }

  static isValidMode(mode: string): mode is DetailOneMode {
    return VALID_MODE_SET.has(mode as DetailOneMode);
  }

  static async getOrInitialize(companyId: string): Promise<DetailSettingsRecord> {
    await this.ensureTableExists();

    const existing = await this.pool.query(
      'SELECT detail_one_mode, updated_at FROM company_detail_settings WHERE company_id = $1',
      [companyId]
    );

    if (existing.rowCount && existing.rows[0]) {
      const row = existing.rows[0];
      return {
        detailOneMode: (row.detail_one_mode || 'auto') as DetailOneMode,
        updatedAt: row.updated_at ? new Date(row.updated_at) : new Date(),
        persisted: true
      };
    }

    const inserted = await this.pool.query(
      `INSERT INTO company_detail_settings (company_id, detail_one_mode)
       VALUES ($1, 'auto')
       ON CONFLICT (company_id) DO UPDATE SET detail_one_mode = EXCLUDED.detail_one_mode
       RETURNING detail_one_mode, updated_at`,
      [companyId]
    );

    const row = inserted.rows[0];
    return {
      detailOneMode: (row.detail_one_mode || 'auto') as DetailOneMode,
      updatedAt: row.updated_at ? new Date(row.updated_at) : new Date(),
      persisted: true
    };
  }

  static async update(companyId: string, mode: DetailOneMode): Promise<DetailSettingsRecord> {
    await this.ensureTableExists();

    const result = await this.pool.query(
      `INSERT INTO company_detail_settings (company_id, detail_one_mode, updated_at)
       VALUES ($1, $2, CURRENT_TIMESTAMP)
       ON CONFLICT (company_id) DO UPDATE SET
         detail_one_mode = EXCLUDED.detail_one_mode,
         updated_at = CURRENT_TIMESTAMP
       RETURNING detail_one_mode, updated_at`,
      [companyId, mode]
    );

    const row = result.rows[0];
    return {
      detailOneMode: (row.detail_one_mode || mode) as DetailOneMode,
      updatedAt: row.updated_at ? new Date(row.updated_at) : new Date(),
      persisted: true
    };
  }
}
