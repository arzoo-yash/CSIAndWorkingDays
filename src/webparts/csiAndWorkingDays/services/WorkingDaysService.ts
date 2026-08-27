import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { LIST_NAMES } from '../../../common/constants';

interface IWorkingDaysCreatePayload {
  Title?: string; // Month field - internal name is Title in SharePoint
  Year?: number;
  Days?: number; // Total working days
  ResourceAllocation?: number;
}

export class WorkingDaysService {
  private spService: SharePointService;

  constructor(context: WebPartContext) {
    this.spService = new SharePointService(context);
  }

  public async getAll(): Promise<any[]> {
    const select = [
      'Id',
      'Title', //internal column name of Month field
      'Year',
      'Days',
      'ResourceAllocation',
      'Modified'
    ];

    // Return items ordered by Year desc, then Month desc
    const items = await this.spService.getListItems(
      LIST_NAMES.Monthly_Workdays, 
      select, 
      undefined, 
      undefined, 
      500, 
      'Modified', 
      false
    );
    return items || [];
  }

  public async getById(id: number): Promise<any> {
    return await this.spService.getListItemById(LIST_NAMES.Monthly_Workdays, id);
  }

  public async create(payload: IWorkingDaysCreatePayload): Promise<any> {
    return await this.spService.createListItem(LIST_NAMES.Monthly_Workdays, payload);
  }

  public async update(id: number, payload: IWorkingDaysCreatePayload): Promise<any> {
    return await this.spService.updateListItem(LIST_NAMES.Monthly_Workdays, id, payload);
  }
}

export default WorkingDaysService;
