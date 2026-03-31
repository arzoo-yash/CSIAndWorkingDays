import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { LIST_NAMES } from '../../../common/constants';

interface ICSATCreatePayload {
  Title?: string; //CSATValue will be saved in Title
  CSATAquiredDate?: string | Date;
  Goal?: string;
  USL?: number;
  LSL?: number;
  Remarks?: string;
}

export class CSATService {
  private spService: SharePointService;
  private context: WebPartContext;
  private rootSiteUrl: string = '';

  constructor(context: WebPartContext) {
    this.context = context;
    this.spService = new SharePointService(context);
  }

  /**
   * Get the root site URL from the current context
   * This extracts the root site from the absolute URL
   * Similar logic to PPOApproverService
   */
  private async getRootSiteUrl(): Promise<string> {
    if (this.rootSiteUrl) {
      return this.rootSiteUrl;
    }

    try {
      // Get the current site absolute URL
      const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
      
      // Parse to get the root site URL
      // For example: https://tenant.sharepoint.com/sites/MySite/SubSite
      // Root would be: https://tenant.sharepoint.com/sites/MySite
      const url = new URL(currentSiteUrl);
      const pathSegments = url.pathname.split('/').filter(s => s);
      
      // Typically, root site structure is /sites/[sitename]
      if (pathSegments.length >= 2 && pathSegments[0] === 'sites') {
        this.rootSiteUrl = `${url.origin}/${pathSegments[0]}/${pathSegments[1]}`;
      } else {
        // If not in /sites/, use the origin as root
        this.rootSiteUrl = url.origin;
      }

      return this.rootSiteUrl;
    } catch (error) {
      console.error('[CSATService] Error determining root site URL:', error);
      // Fallback to current site
      return this.context.pageContext.web.absoluteUrl;
    }
  }

  /**
   * Check if the current site is the root site
   */
  private async isRootSite(): Promise<boolean> {
    try {
      const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
      const rootSiteUrl = await this.getRootSiteUrl();
      return currentSiteUrl === rootSiteUrl;
    } catch (error) {
      console.error('[CSATService] Error checking if current site is root:', error);
      return false;
    }
  }

  /**
   * Sync CSI entry to root site
   * Creates or updates the entry in the root site's CSI list
   */
  private async syncToRootSite(payload: ICSATCreatePayload, currentSiteItemId: number): Promise<void> {
    try {
      // Check if we're already at the root site - if so, no sync needed
      if (await this.isRootSite()) {
        console.debug('[CSATService] Already at root site, skipping sync to root');
        return;
      }

      const rootSiteUrl = await this.getRootSiteUrl();
      const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
      console.debug('[CSATService] Syncing CSI entry to root site:', rootSiteUrl);

      // Prepare payload for root site (same fields + ProjectSiteURL)
      const rootPayload: any = {
        Title: payload.Title,
        CSATAquiredDate: payload.CSATAquiredDate,
        Goal: payload.Goal,
        USL: payload.USL,
        LSL: payload.LSL,
        Remarks: payload.Remarks,
        ProjectSiteURL: currentSiteUrl // Store the subsite URL where this entry was created
      };

      // Try to find existing entry in root site by CSATAquiredDate and ProjectSiteURL
      // This combination should uniquely identify entries from the same subsite
      let rootItemId: number | null = null;

      if (payload.CSATAquiredDate) {
        try {
          const dateFilter = this.buildDateFilter(payload.CSATAquiredDate);
          const siteFilter = `ProjectSiteURL eq '${currentSiteUrl}'`;
          const combinedFilter = `${dateFilter} and ${siteFilter}`;
          const existingItems = await this.spService.getListItemsFromSite(
            rootSiteUrl,
            LIST_NAMES.Customer_Satisfaction_Index,
            ['Id', 'Title', 'CSATAquiredDate', 'ProjectSiteURL'],
            undefined,
            combinedFilter,
            1
          );

          if (existingItems && existingItems.length > 0) {
            rootItemId = existingItems[0].Id;
            console.debug('[CSATService] Found existing root CSI item:', rootItemId);
          }
        } catch (err) {
          console.warn('[CSATService] Could not search for existing root CSI item:', err);
        }
      }

      // Create or update in root site
      if (rootItemId) {
        // Update existing entry in root site
        await this.spService.updateListItemInSite(
          rootSiteUrl,
          LIST_NAMES.Customer_Satisfaction_Index,
          rootItemId,
          rootPayload
        );
        console.debug('[CSATService] Updated root CSI item:', rootItemId);
      } else {
        // Create new entry in root site
        const createdItem = await this.spService.createListItemInSite(
          rootSiteUrl,
          LIST_NAMES.Customer_Satisfaction_Index,
          rootPayload
        );
        rootItemId = createdItem?.Id || createdItem?.data?.Id;
        console.debug('[CSATService] Created new root CSI item:', rootItemId);
      }
    } catch (error) {
      // Log error but don't fail the main operation
      console.error('[CSATService] Error syncing to root site:', error);
      // We don't throw here to ensure the main create/update operation succeeds
    }
  }

  /**
   * Build an OData filter for date comparison
   * Handles both Date objects and ISO strings
   */
  private buildDateFilter(date: string | Date): string {
    const dateObj = date instanceof Date ? date : new Date(date);
    const isoDate = dateObj.toISOString();
    // Filter by date only (ignoring time component)
    const dateStr = isoDate.split('T')[0];
    return `startswith(CSATAquiredDate, '${dateStr}')`;
  }

  public async getAll(): Promise<any[]> {
    // Select common fields we expect to use in the UI
    const select = [
      'Id',
      'Title',
      'CSATAquiredDate',
      'Goal',
      'USL',
      'LSL',
      'Remarks',
      'Modified'
    ];

    // Return items ordered by Modified desc (SharePointService defaults to Modified desc)
    const items = await this.spService.getListItems(LIST_NAMES.Customer_Satisfaction_Index, select, undefined, undefined, 500, 'Modified', false);
    return items || [];
  }

  public async getById(id: number): Promise<any> {
    const select = [
      'Id',
      'Title',
      'CSATAquiredDate',
      'Goal',
      'USL',
      'LSL',
      'Remarks',
      'Modified',
      'Created'
    ];
    return await this.spService.getListItemById(LIST_NAMES.Customer_Satisfaction_Index, id, select);
  }

  public async create(payload: ICSATCreatePayload): Promise<any> {
    // Create item in current site first
    const result = await this.spService.createListItem(LIST_NAMES.Customer_Satisfaction_Index, payload);
    
    // Sync to root site asynchronously (non-blocking)
    const itemId = result?.Id || result?.data?.Id;
    if (itemId) {
      // Don't await - let it run in background
      void this.syncToRootSite(payload, itemId);
    }
    
    return result;
  }

  public async update(id: number, payload: ICSATCreatePayload): Promise<any> {
    // Update item in current site first
    const result = await this.spService.updateListItem(LIST_NAMES.Customer_Satisfaction_Index, id, payload);
    
    // Sync to root site asynchronously (non-blocking)
    // Don't await - let it run in background
    void this.syncToRootSite(payload, id);
    
    return result;
  }
}

export default CSATService;
