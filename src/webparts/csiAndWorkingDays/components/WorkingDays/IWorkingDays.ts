import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { DashboardTabKey } from '../types';

export interface IWorkingDaysRecord {
  id: number;
  month?: string;
  year?: number;
  periodLabel: string;
  totalDays?: number;
  resourceAllocation?: number;
  billableDays?: number;
  leaves?: number;
  holidays?: number;
  remarks?: string;
  modified?: Date;
}

export interface IWorkingDaysProps {
  context: WebPartContext;
  onSelectTab: (tab: DashboardTabKey) => void;
}

export interface IWorkingDaysFormData {
  Title?: string; // Month field - internal name is Title in SharePoint
  Year?: number;
  Days?: number;
  ResourceAllocation?: number;
}

export interface IWorkingDaysFormProps {
  context: WebPartContext;
  item?: any;
  isEdit?: boolean;
  onSaved?: (result: any) => void;
  onCancel?: () => void;
}
