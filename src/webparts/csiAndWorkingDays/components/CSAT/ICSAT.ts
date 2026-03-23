import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { DashboardTabKey } from '../types';

export interface ICSATRecord {
  id: number;
  csatValue?: string; // Title field
  remarks?: string;
  csatAquiredDate?: Date;
  goal?: string;
  usl?: number;
  lsl?: number;
  modified?: Date;
}

export interface ICSATProps {
  context: WebPartContext;
  onSelectTab: (tab: DashboardTabKey) => void;
}
