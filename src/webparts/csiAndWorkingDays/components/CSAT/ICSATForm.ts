import type { WebPartContext } from '@microsoft/sp-webpart-base';

export interface INewEditCSATFormData {
  Title?: string;
  CSATAquiredDate?: string | Date;
  Goal?: string;
  USL?: number;
  LSL?: number;
  Remarks?: string;
}

export interface INewEditCSATProps {
  context: WebPartContext;
  item?: any;
  isEdit?: boolean;
  onSaved?: (result: any) => void;
  onCancel?: () => void;
}
