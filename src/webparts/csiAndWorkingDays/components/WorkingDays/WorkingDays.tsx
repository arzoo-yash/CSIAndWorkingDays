import * as React from 'react';
import {
  MessageBar,
  MessageBarType,
  Spinner,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Dialog,
  DialogType,
  IconButton,
  PrimaryButton,
  DefaultButton
} from '@fluentui/react';
import styles from './WorkingDays.module.scss';
import headerStyles from '../Header/Header.module.scss';
import type { IWorkingDaysProps, IWorkingDaysRecord } from './IWorkingDays';
import WorkingDaysService from '../../services/WorkingDaysService';
import { PPOApproverService } from '../../services/PPOApproverService';
import Header from '../Header';
import type { DashboardTabKey } from '../types';
import WorkingDaysForm from './WorkingDaysForm';

const PAGE_SIZE = 8;
const MONTH_ORDER: Record<string, number> = {
  january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
  july: 7, august: 8, september: 9, october: 10, november: 11, december: 12
};

type SortableColumnKey = 'year' | 'month';

const WorkingDays: React.FC<IWorkingDaysProps> = ({ context, onSelectTab }) => {
  const [records, setRecords] = React.useState<IWorkingDaysRecord[]>([]);
  const [rawItems, setRawItems] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [editingItem, setEditingItem] = React.useState<any | undefined>(undefined);
  const [canEdit, setCanEdit] = React.useState<boolean>(false);
  const [sortColumnKey, setSortColumnKey] = React.useState<SortableColumnKey | undefined>(undefined);
  const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);
  const [currentPage, setCurrentPage] = React.useState<number>(1);

  const service = React.useMemo(() => new WorkingDaysService(context), [context]);
  const ppoService = React.useMemo(() => new PPOApproverService(context), [context]);

  const pickString = React.useCallback((item: any, keys: string[]): string | undefined => {
    for (const key of keys) {
      const value = item?.[key];
      if (typeof value === 'string' && value.trim().length > 0) {
        return value.trim();
      }
    }
    return undefined;
  }, []);

  const pickNumber = React.useCallback((item: any, keys: string[]): number | undefined => {
    for (const key of keys) {
      const value = item?.[key];
      if (typeof value === 'number' && !isNaN(value)) {
        return value;
      }
      if (typeof value === 'string') {
        const parsed = parseFloat(value);
        if (!isNaN(parsed)) {
          return parsed;
        }
      }
    }
    return undefined;
  }, []);

  const mapRecords = React.useCallback((items: any[]): IWorkingDaysRecord[] => items.map((item: any) => {
    const month = pickString(item, ['Title', 'Month', 'ReportingMonth', 'Period']);
    const year = pickNumber(item, ['Year', 'ReportingYear']);
    const labelParts = [month, year].filter(Boolean) as (string | number)[];
    return {
      id: item?.Id ?? item?.ID ?? 0,
      month: month,
      year: year,
      periodLabel: labelParts.length > 0 ? labelParts.join(' ') : month ?? 'Not specified',
      totalDays: pickNumber(item, ['Days', 'TotalWorkingDays', 'WorkingDays', 'TotalDays']),
      resourceAllocation: pickNumber(item, ['ResourceAllocation']),
      billableDays: pickNumber(item, ['BillableDays', 'Billable']),
      leaves: pickNumber(item, ['Leaves', 'PlannedLeaves', 'ApprovedLeaves']),
      holidays: pickNumber(item, ['Holidays', 'NonWorkingDays']),
      remarks: pickString(item, ['Remarks', 'Notes', 'Comments']),
      modified: item?.Modified ? new Date(item.Modified) : undefined
    };
  }), [pickNumber, pickString]);

  const loadData = React.useCallback(async () => {
    setLoading(true);
    try {
      const items = await service.getAll();
      setRawItems(items || []);
      setRecords(mapRecords(items || []));
      setError(null);
    } catch (err) {
      console.error('[WorkingDays] Failed to load data', err);
      setError('We could not load the latest working days entries. Please try again later.');
    } finally {
      setLoading(false);
    }
  }, [mapRecords, service]);

  React.useEffect(() => {
    void loadData();
  }, [loadData]);

  const sortedRecords = React.useMemo(() => {
    if (!sortColumnKey) {
      return records;
    }
    return [...records].sort((a, b) => {
      let compareResult = 0;
      if (sortColumnKey === 'year') {
        compareResult = (a.year ?? 0) - (b.year ?? 0);
      } else {
        const aOrder = MONTH_ORDER[(a.month ?? '').toLowerCase()] ?? 0;
        const bOrder = MONTH_ORDER[(b.month ?? '').toLowerCase()] ?? 0;
        compareResult = aOrder - bOrder;
      }
      return isSortedDescending ? -compareResult : compareResult;
    });
  }, [records, sortColumnKey, isSortedDescending]);

  const totalPages = Math.max(1, Math.ceil(sortedRecords.length / PAGE_SIZE));

  const pagedRecords = React.useMemo(() => {
    const start = (currentPage - 1) * PAGE_SIZE;
    return sortedRecords.slice(start, start + PAGE_SIZE);
  }, [sortedRecords, currentPage]);

  React.useEffect(() => {
    setCurrentPage(1);
  }, [records, sortColumnKey, isSortedDescending]);

  React.useEffect(() => {
    setCurrentPage(prev => Math.min(prev, totalPages));
  }, [totalPages]);

  const onColumnHeaderClick = React.useCallback((key: SortableColumnKey) => {
    setIsSortedDescending(sortColumnKey === key ? !isSortedDescending : false);
    setSortColumnKey(key);
  }, [sortColumnKey, isSortedDescending]);

  React.useEffect(() => {
    const checkPermissions = async (): Promise<void> => {
      try {
        const hasEditPermission = await ppoService.canUserEdit();
        setCanEdit(hasEditPermission);
      } catch (err) {
        console.error('[WorkingDays] Failed to check permissions', err);
        setCanEdit(false);
      }
    };
    void checkPermissions();
  }, [ppoService]);

  const formatNumber = React.useCallback((value?: number, fractionDigits: number = 0): string => {
    if (typeof value === 'number' && !isNaN(value)) {
      return value.toFixed(fractionDigits);
    }
    return '—';
  }, []);

  const handleTabChange = React.useCallback((tabKey: string) => {
    onSelectTab(tabKey as DashboardTabKey);
  }, [onSelectTab]);

  const openNew = React.useCallback(() => {
    setEditingItem(undefined);
    setIsPanelOpen(true);
  }, []);

  const openEdit = React.useCallback((item: any) => {
    // Find the original raw item by ID
    let rawItem: any = item;
    for (let i = 0; i < rawItems.length; i++) {
      if ((rawItems[i].Id || rawItems[i].ID) === item.id) {
        rawItem = rawItems[i];
        break;
      }
    }
    setEditingItem(rawItem);
    setIsPanelOpen(true);
  }, [rawItems]);

  const closePanel = React.useCallback(() => {
    setIsPanelOpen(false);
    setEditingItem(undefined);
  }, []);

  const columns = React.useMemo<IColumn[]>(() => [
    { 
      key: 'colActions', 
      name: 'Actions', 
      fieldName: 'actions', 
      minWidth: 80, 
      maxWidth: 80,
      onRender: (item: any) => (
        canEdit ? (
          <IconButton 
            title="Edit" 
            ariaLabel="Edit" 
            iconProps={{ iconName: 'Edit' }} 
            onClick={() => openEdit(item)} 
          />
        ) : null
      ) 
    },
    {
      key: 'colYear',
      name: 'Year',
      fieldName: 'year',
      minWidth: 100,
      maxWidth: 100,
      isSorted: sortColumnKey === 'year',
      isSortedDescending: sortColumnKey === 'year' && isSortedDescending,
      onColumnClick: () => onColumnHeaderClick('year')
    },
    {
      key: 'colMonth',
      name: 'Month',
      fieldName: 'month',
      minWidth: 120,
      maxWidth: 120,
      isSorted: sortColumnKey === 'month',
      isSortedDescending: sortColumnKey === 'month' && isSortedDescending,
      onColumnClick: () => onColumnHeaderClick('month')
    },
    { key: 'colWorkingDays', name: 'Working Days', fieldName: 'totalDays', minWidth: 140, maxWidth: 140, onRender: (item: IWorkingDaysRecord) => formatNumber(item.totalDays) },
    { key: 'colResourceAllocation', name: 'Allocated Resources', fieldName: 'resourceAllocation', minWidth: 160, maxWidth: 160, onRender: (item: IWorkingDaysRecord) => formatNumber(item.resourceAllocation, 2) }
  ], [formatNumber, openEdit, canEdit, rawItems, sortColumnKey, isSortedDescending, onColumnHeaderClick]);

  const tabs = React.useMemo(() => [
    { key: 'csat', label: 'Customer Satisfaction' }, 
    { key: 'workingDays', label: 'No.Working Days' }
  ], []);

  return (
    <section className={styles.tabShell} aria-labelledby="working-days-heading">
      <Header
        title="Admin Lists"
        iconName="Bullseye"
      />

      <div className={styles.tabButtons}>
        {tabs.map(tab => (
          <button
            key={tab.key}
            type="button"
            className={`${styles.tabButton} ${tab.key === 'workingDays' ? styles.activeTab : ''}`.trim()}
            onClick={() => handleTabChange(tab.key)}
          >
            {tab.label}
          </button>
        ))}
      </div>

      <div className={styles.actionRow}>
        {canEdit && (
          <PrimaryButton
            className={headerStyles.primaryAction}
            onClick={openNew}
            ariaLabel="Add new"
          >
            <span>+ Add New</span>
          </PrimaryButton>
        )}
      </div>

      <div className={styles.panel}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} className={styles.messageBar}>
            {error}
          </MessageBar>
        )}

        {loading ? (
          <div className={styles.loadingState}>
            <Spinner label="Loading working days data" />
          </div>
        ) : (
          <>
            <div className={styles.tableWrapper}>
              {sortedRecords.length > 0 ? (
                <DetailsList
                  items={pagedRecords}
                  columns={columns}
                  setKey="workingDays"
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  selectionMode={SelectionMode.none}
                />
              ) : (
                <div className={styles.emptyState}>No working days entries available.</div>
              )}
            </div>

            {sortedRecords.length > 0 && (
              <div className={styles.paginationBar}>
                <span className={styles.paginationInfo}>Page {currentPage} of {totalPages}</span>
                <div className={styles.paginationButtons}>
                  <DefaultButton
                    text="Previous"
                    disabled={currentPage <= 1}
                    onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                  />
                  <PrimaryButton
                    text="Next"
                    disabled={currentPage >= totalPages}
                    onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                  />
                </div>
              </div>
            )}
          </>
        )}

        <Dialog
          hidden={!isPanelOpen}
          onDismiss={closePanel}
          dialogContentProps={{
            type: DialogType.normal,
            title: editingItem ? 'Edit Working Days' : 'New Working Days Entry'
          }}
          minWidth={500}
          maxWidth={600}
        >
          <WorkingDaysForm
            context={context}
            item={editingItem}
            isEdit={!!editingItem}
            onSaved={() => { void loadData(); closePanel(); }}
            onCancel={closePanel}
          />
        </Dialog>
      </div>
    </section>
  );
};

export default WorkingDays;
