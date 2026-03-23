import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/batching';
import '@pnp/sp/fields';

export interface ISharePointService {
  getListItems(listName: string, select?: string[], expand?: string[], filter?: string, top?: number, orderBy?: string, ascending?: boolean): Promise<any[]>;
  getListItemsFromSite(siteUrl: string, listName: string, select?: string[], expand?: string[], filter?: string, top?: number, orderBy?: string, ascending?: boolean): Promise<any[]>;
  getListItemById(listName: string, itemId: number, select?: string[], expand?: string[]): Promise<any>;
  createListItem(listName: string, item: any): Promise<any>;
  updateListItem(listName: string, itemId: number, item: any): Promise<any>;
  deleteListItem(listName: string, itemId: number): Promise<void>;
  getListItemsByCAML(listName: string, camlQuery: string): Promise<any[]>;
  getItemVersionHistory(listName: string, itemId: number): Promise<any[]>;
  ensureUsers(loginNames: string[]): Promise<number[]>;
  getPeopleField(listName: string, itemId: number, fieldInternalName: string): Promise<any[]>;
  updatePeopleFieldByIds(listName: string, itemId: number, fieldInternalName: string, userIds: number[]): Promise<any>;
  createListItemWithPeople(listName: string, item: any, fieldInternalName: string, userIds: number[]): Promise<any>;
  createListItemWithPeopleRaw(listName: string, item: any, fieldInternalName: string, userIds: number[]): Promise<any>;
}

export class SharePointService implements ISharePointService {
  private sp: SPFI;
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  // Ensure fields that allow multiple values receive an array/envelope shape
  private async sanitizePayloadForMultiValue(listName: string, payload: any): Promise<void> {
    if (!payload || typeof payload !== 'object') return;
    const keys = Object.keys(payload);
    for (const key of keys) {
      try {
        // Skip internal peopleId keys (they end with 'Id') as we handle them separately
        if (key.slice(-2) === 'Id') continue;
        const allowMulti = await this.getFieldAllowMultiple(listName, key);
        if (allowMulti) {
          const val = payload[key];
          // If it's an empty string or primitive where an array is expected, provide an empty envelope
          if (val === '' || val === null || typeof val === 'number' || typeof val === 'boolean') {
            payload[key] = { results: [] };
          }
          // If it's an array, leave as-is (caller code will convert to envelope if necessary)
        }
      } catch (err) {
        // If reading field metadata fails, don't block creation; continue
        // eslint-disable-next-line no-console
        console.debug('[SharePointService] sanitizePayloadForMultiValue - skipped field', { listName, field: key, err });
      }
    }
  }

  /**
   * Ensure users exist in the site and return their Ids.
   * Accepts login names or emails (e.g. 'i:0#.f|membership|user@domain.com' or 'user@domain.com').
   */
  public async ensureUsers(loginNames: string[]): Promise<number[]> {
    try {
      const ids: number[] = [];
      for (const login of loginNames) {
        const res: any = await this.sp.web.ensureUser(login);
        // depending on PnP version the id may be under data or returned directly
        const id = res && (res.data?.Id ?? res.Id ?? res.id);
        if (!id) {
          throw new Error(`Unable to resolve user '${login}'`);
        }
        ids.push(id as number);
      }
      return ids;
    } catch (error) {
      console.error('Error ensuring users:', error);
      throw error;
    }
  }

  /**
   * Read a multi-user People Picker field from a list item and return an array of user objects.
   */
  public async getPeopleField(listName: string, itemId: number, fieldInternalName: string): Promise<any[]> {
    try {
      const selectFields = [`${fieldInternalName}/Id`, `${fieldInternalName}/Title`, `${fieldInternalName}/EMail`];
      const item = await this.sp.web.lists.getByTitle(listName).items.getById(itemId).select(...selectFields).expand(fieldInternalName)();

      const fieldValue = item ? item[fieldInternalName] : null;
      if (!fieldValue) {
        return [];
      }

      // fieldValue may be an array (multi user) or a single object
      if (Array.isArray(fieldValue)) {
        return fieldValue.map((u: any) => ({ Id: u.Id, Title: u.Title, Email: u.EMail || u.Email }));
      }

      return [{ Id: fieldValue.Id, Title: fieldValue.Title, Email: fieldValue.EMail || fieldValue.Email }];
    } catch (error) {
      console.error(`Error reading people field ${fieldInternalName} from item ${itemId} in list ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Update a multi-user People Picker field by passing an array of user IDs.
   * Use the internal field name without the trailing 'Id' (we append 'Id' automatically).
   */
  public async updatePeopleFieldByIds(listName: string, itemId: number, fieldInternalName: string, userIds: number[]): Promise<any> {
    try {
      const payload: any = {};

      // Determine if the target people field allows multiple values
      const field = await this.sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(fieldInternalName)();
      const allowMulti = !!(field && (field as any).AllowMultipleValues);

      if (allowMulti) {
        // Some tenants prefer the plain array form (alternate), others accept the { results: [...] } envelope.
        // Try the plain array first and fall back to the envelope form on a payload-shape error.
        const altPayload: any = {};
        altPayload[`${fieldInternalName}Id`] = (!userIds || userIds.length === 0) ? [] : userIds;

        try {
          await this.sp.web.lists.getByTitle(listName).items.getById(itemId).update(altPayload);
        } catch (err: any) {
          const msg = (err && err.message) ? err.message : (err && err.responseText) ? err.responseText : '';
          // If SharePoint complained about JSON node types, retry with the envelope form
          if (msg && (msg.indexOf('StartObject') !== -1 || msg.indexOf('StartArray') !== -1 || msg.indexOf('InvalidClientQueryException') !== -1)) {
            const envelope: any = {};
            envelope[`${fieldInternalName}Id`] = (!userIds || userIds.length === 0) ? { results: [] } : { results: userIds };
            await this.sp.web.lists.getByTitle(listName).items.getById(itemId).update(envelope);
          } else {
            throw err;
          }
        }
      } else {
        // Single-user field expects a single number or null to clear
        if (!userIds || userIds.length === 0) {
          payload[`${fieldInternalName}Id`] = null;
        } else {
          payload[`${fieldInternalName}Id`] = userIds[0];
        }

        await this.sp.web.lists.getByTitle(listName).items.getById(itemId).update(payload);
      }
      return await this.getListItemById(listName, itemId, [
        `${fieldInternalName}/Id`,
        `${fieldInternalName}/Title`,
        `${fieldInternalName}/EMail`
      ], [fieldInternalName]);
    } catch (error) {
      console.error(`Error updating people field ${fieldInternalName} on item ${itemId} in list ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Create a new list item and set a multi-user People Picker field by user IDs.
   */
  public async createListItemWithPeople(listName: string, item: any, fieldInternalName: string, userIds: number[]): Promise<any> {
    try {
      const payload = { ...item };
      // Sanitize payload for multi-value fields
      try { await this.sanitizePayloadForMultiValue(listName, payload); } catch (e) { /* ignore */ }
      
      const field = await this.sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(fieldInternalName)();
      const allowMulti = !!(field && (field as any).AllowMultipleValues);

      if (allowMulti) {
        // Try plain array first (works for most standard SharePoint configurations)
        payload[`${fieldInternalName}Id`] = userIds.length > 0 ? userIds : [];
        
        try {
          const result = await this.sp.web.lists.getByTitle(listName).items.add(payload);
          return this.extractResult(result);
        } catch (err: any) {
          const msg = (err && err.message) ? err.message : (err && err.responseText) ? err.responseText : '';
          
          // If plain array fails, try envelope format
          if (msg && (msg.indexOf('StartObject') !== -1 || msg.indexOf('StartArray') !== -1 || msg.indexOf('InvalidClientQueryException') !== -1)) {
            const envelopePayload = { ...item };
            // Re-sanitize the envelope payload
            try { await this.sanitizePayloadForMultiValue(listName, envelopePayload); } catch (e) { /* ignore */ }
            envelopePayload[`${fieldInternalName}Id`] = userIds.length > 0 ? { results: userIds } : { results: [] };
            const result2 = await this.sp.web.lists.getByTitle(listName).items.add(envelopePayload);
            return this.extractResult(result2);
          }
          throw err;
        }
      } else {
        payload[`${fieldInternalName}Id`] = userIds.length > 0 ? userIds[0] : null;
        const result = await this.sp.web.lists.getByTitle(listName).items.add(payload);
        return this.extractResult(result);
      }
    } catch (error) {
      console.error(`Error creating item in list ${listName} with people field ${fieldInternalName}:`, error);
      throw error;
    }
  }

  /**
   * Create a new list item with people field without any payload sanitization.
   * Use this for complex payloads that already have properly formatted multi-value fields.
   */
  public async createListItemWithPeopleRaw(listName: string, item: any, fieldInternalName: string, userIds: number[]): Promise<any> {
    try {
      const payload = { ...item };
      const field = await this.sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(fieldInternalName)();
      const allowMulti = !!(field && (field as any).AllowMultipleValues);

      // For item creation, use plain array for multi-value fields (not envelope format)
      if (allowMulti) {
        payload[`${fieldInternalName}Id`] = userIds.length > 0 ? userIds : [];
      } else {
        payload[`${fieldInternalName}Id`] = userIds.length > 0 ? userIds[0] : null;
      }

      const result = await this.sp.web.lists.getByTitle(listName).items.add(payload);
      return this.extractResult(result);
    } catch (error) {
      console.error(`Error creating item in list ${listName} with people field ${fieldInternalName}:`, error);
      throw error;
    }
  }

  /**
   * Helper to extract result from PnP response
   */
  private extractResult(result: any): any {
    if (result && (result as any).data) return (result as any).data;
    if (result && (result as any).item && typeof (result as any).item.get === 'function') {
      return (result as any).item.get();
    }
    return result;
  }

  /**
   * Get all items from a SharePoint list
   * @param listName - Name of the SharePoint list
   * @param select - Array of fields to select
   * @param expand - Array of fields to expand
   * @param filter - OData filter string
   * @param top - Number of items to retrieve
   * @param orderBy - Field name to order by
   * @param ascending - Order direction (true for ascending, false for descending)
   * @returns Promise with array of list items
   */
  public async getListItems(
    listName: string,
    select?: string[],
    expand?: string[],
    filter?: string,
    top?: number,
    orderBy?: string,
    ascending?: boolean
  ): Promise<any[]> {
    try {
      let query = this.sp.web.lists.getByTitle(listName).items;

      if (select && select.length > 0) {
        query = query.select(...select);
      }

      if (expand && expand.length > 0) {
        query = query.expand(...expand);
      }

      if (filter) {
        query = query.filter(filter);
      }

      if (top) {
        query = query.top(top);
      }

      // Apply ordering - use custom orderBy if provided, otherwise default to Modified descending
      if (orderBy) {
        query = query.orderBy(orderBy, ascending !== false);
      } else {
        query = query.orderBy('Modified', false);
      }

      const items = await query();
      return items;
    } catch (error) {
      console.error(`Error getting items from list ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Get items from a SharePoint list in a different site (e.g., root site)
   * @param siteUrl - The absolute URL of the site to query
   * @param listName - Name of the SharePoint list
   * @param select - Array of fields to select
   * @param expand - Array of fields to expand
   * @param filter - OData filter string
   * @param top - Number of items to retrieve
   * @param orderBy - Field name to order by
   * @param ascending - Order direction (true for ascending, false for descending)
   * @returns Promise with array of list items
   */
  public async getListItemsFromSite(
    siteUrl: string,
    listName: string,
    select?: string[],
    expand?: string[],
    filter?: string,
    top?: number,
    orderBy?: string,
    ascending?: boolean
  ): Promise<any[]> {
    try {
      // Create a new SPFI instance targeting the specific site URL
      const targetSp = spfi(siteUrl).using(SPFx(this.context));
      let query = targetSp.web.lists.getByTitle(listName).items;

      if (select && select.length > 0) {
        query = query.select(...select);
      }

      if (expand && expand.length > 0) {
        query = query.expand(...expand);
      }

      if (filter) {
        query = query.filter(filter);
      }

      if (top) {
        query = query.top(top);
      }

      if (orderBy) {
        query = query.orderBy(orderBy, ascending !== false);
      } else {
        query = query.orderBy('Modified', false);
      }

      const items = await query();
      return items;
    } catch (error) {
      console.error(`Error getting items from list ${listName} in site ${siteUrl}:`, error);
      throw error;
    }
  }

  /**
   * Get a specific item by ID from a SharePoint list
   * @param listName - Name of the SharePoint list
   * @param itemId - ID of the item to retrieve
   * @param select - Array of fields to select
   * @param expand - Array of fields to expand
   * @returns Promise with the list item
   */
  public async getListItemById(
    listName: string,
    itemId: number,
    select?: string[],
    expand?: string[]
  ): Promise<any> {
    try {
      let query = this.sp.web.lists.getByTitle(listName).items.getById(itemId);

      if (select && select.length > 0) {
        query = query.select(...select);
      }

      if (expand && expand.length > 0) {
        query = query.expand(...expand);
      }

      const item = await query();
      return item;
    } catch (error) {
      console.error(`Error getting item ${itemId} from list ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Create a new item in a SharePoint list
   * @param listName - Name of the SharePoint list
   * @param item - Object containing the item data
   * @returns Promise with the created item
   */
  public async createListItem(listName: string, item: any): Promise<any> {
    try {
      // sanitize payload for multi-value fields
      const payload = { ...item };
      try { await this.sanitizePayloadForMultiValue(listName, payload); } catch (e) { /* ignore */ }
      // Debug: log payload being sent to create
      // eslint-disable-next-line no-console
      console.debug('[SharePointService] createListItem payload:', { listName, payload });
      const result = await this.sp.web.lists.getByTitle(listName).items.add(payload);
      if (result && (result as any).data) return (result as any).data;
      if (result && (result as any).item && typeof (result as any).item.get === 'function') {
        return await (result as any).item.get();
      }
      return result;
    } catch (error) {
      console.error(`Error creating item in list ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Create a new item in a SharePoint list at a specific site
   * @param siteUrl - The absolute URL of the site
   * @param listName - Name of the SharePoint list
   * @param item - Object containing the item data
   * @returns Promise with the created item
   */
  public async createListItemInSite(siteUrl: string, listName: string, item: any): Promise<any> {
    try {
      // Create a new SPFI instance targeting the specific site URL
      const targetSp = spfi(siteUrl).using(SPFx(this.context));
      
      // Debug: log payload being sent to create
      // eslint-disable-next-line no-console
      console.debug('[SharePointService] createListItemInSite payload:', { siteUrl, listName, item });
      
      const result = await targetSp.web.lists.getByTitle(listName).items.add(item);
      if (result && (result as any).data) return (result as any).data;
      if (result && (result as any).item && typeof (result as any).item.get === 'function') {
        return await (result as any).item.get();
      }
      return result;
    } catch (error) {
      console.error(`Error creating item in list ${listName} at site ${siteUrl}:`, error);
      throw error;
    }
  }

  /**
   * Update an existing item in a SharePoint list
   * @param listName - Name of the SharePoint list
   * @param itemId - ID of the item to update
   * @param item - Object containing the updated data
   * @returns Promise with the updated item
   */
  public async updateListItem(listName: string, itemId: number, item: any): Promise<any> {
    try {
      // Debug: log payload being sent to update
      // eslint-disable-next-line no-console
      console.debug('[SharePointService] updateListItem payload:', { listName, itemId, item });
      await this.sp.web.lists.getByTitle(listName).items.getById(itemId).update(item);
      return await this.getListItemById(listName, itemId);
    } catch (error) {
      console.error(`Error updating item ${itemId} in list ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Delete an item from a SharePoint list
   * @param listName - Name of the SharePoint list
   * @param itemId - ID of the item to delete
   * @returns Promise
   */
  public async deleteListItem(listName: string, itemId: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(listName).items.getById(itemId).delete();
    } catch (error) {
      console.error(`Error deleting item ${itemId} from list ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Get items from a SharePoint list using CAML query
   * @param listName - Name of the SharePoint list
   * @param camlQuery - CAML query string
   * @returns Promise with array of list items
   */
  public async getListItemsByCAML(listName: string, camlQuery: string): Promise<any[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(listName).getItemsByCAMLQuery({
        ViewXml: camlQuery
      });
      return items;
    } catch (error) {
      console.error(`Error getting items from list ${listName} using CAML query:`, error);
      throw error;
    }
  }

  /**
   * Get version history of a list item
   * @param listName - Name of the SharePoint list
   * @param itemId - ID of the item to get version history for
   * @returns Promise with array of item versions
   */
  public async getItemVersionHistory(listName: string, itemId: number): Promise<any[]> {
    try {
      const versions = await this.sp.web.lists.getByTitle(listName).items.getById(itemId).versions();
      return versions;
    } catch (error) {
      console.error(`Error getting version history for item ${itemId} in list ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Returns whether a field allows multiple values (useful for choice and people fields)
   */
  public async getFieldAllowMultiple(listName: string, fieldInternalName: string): Promise<boolean> {
    try {
      const field = await this.sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(fieldInternalName)();
      return !!(field && (field as any).AllowMultipleValues);
    } catch (error) {
      console.error(`Error reading field metadata ${fieldInternalName} from list ${listName}:`, error);
      // default to false if we can't determine
      return false;
    }
  }

  /**
   * Get choices for a Choice field on a list
   * @param listName - Name of the list
   * @param fieldInternalName - Internal name or display name of the field
   * @returns Promise with an array of string choices (empty array on error)
   */
  public async getFieldChoices(listName: string, fieldInternalName: string): Promise<string[]> {
    try {
      const field: any = await this.sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(fieldInternalName)();
      if (!field) return [];
      // Choice fields expose a Choices array
      const choices = field.Choices || field.Choice || [];
      // Ensure it's an array of strings
      if (Array.isArray(choices)) return choices.map((c: any) => String(c));
      return [];
    } catch (error) {
      console.error(`Error getting choices for field ${fieldInternalName} in list ${listName}:`, error);
      return [];
    }
  }

  /**
   * Get the current logged-in user's email address
   * @returns Promise with the current user's email
   */
  public getCurrentUserEmail(): string {
    return this.context.pageContext.user.email || this.context.pageContext.user.loginName;
  }

  /**
   * Get the current logged-in user's login name
   * @returns Promise with the current user's login name
   */
  public getCurrentUserLoginName(): string {
    return this.context.pageContext.user.loginName;
  }
}
