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

  constructor(context: WebPartContext) {
    this.spService = new SharePointService(context);
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
    return await this.spService.createListItem(LIST_NAMES.Customer_Satisfaction_Index, payload);
  }

  public async update(id: number, payload: ICSATCreatePayload): Promise<any> {
    return await this.spService.updateListItem(LIST_NAMES.Customer_Satisfaction_Index, id, payload);
  }
}

export default CSATService;
