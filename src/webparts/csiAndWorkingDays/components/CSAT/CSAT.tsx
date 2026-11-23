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
import styles from './CSAT.module.scss';
import headerStyles from '../Header/Header.module.scss';
import type { ICSATProps, ICSATRecord } from './ICSAT';
import CSATService from '../../services/CSATService';
import { PPOApproverService } from '../../services/PPOApproverService';
import Header from '../Header';
import type { DashboardTabKey } from '../types';
import NewEditCSAT from './CSATForm';

const PAGE_SIZE = 8;

const CSAT: React.FC<ICSATProps> = ({ context, onSelectTab }) => {
  const [records, setRecords] = React.useState<ICSATRecord[]>([]);
  const [rawItems, setRawItems] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [editingItem, setEditingItem] = React.useState<any | undefined>(undefined);
  const [canEdit, setCanEdit] = React.useState<boolean>(false);
  const [isDateSorted, setIsDateSorted] = React.useState<boolean>(false);
  const [isDateSortedDescending, setIsDateSortedDescending] = React.useState<boolean>(false);
  const [currentPage, setCurrentPage] = React.useState<number>(1);

  const service = React.useMemo(() => new CSATService(context), [context]);
  const ppoService = React.useMemo(() => new PPOApproverService(context), [context]);

  const mapRecords = React.useCallback((items: any[]): ICSATRecord[] => items.map((item: any) => ({
    id: item?.Id ?? item?.ID ?? 0,
    csatValue: item?.Title,
    remarks: item?.Remarks,
    csatAquiredDate: item?.CSATAquiredDate ? new Date(item.CSATAquiredDate) : undefined,
    goal: item?.Goal,
    usl: item?.USL,
    lsl: item?.LSL,
    modified: item?.Modified ? new Date(item.Modified) : undefined,
    attachments: Array.isArray(item?.AttachmentFiles)
      ? item.AttachmentFiles.map((f: any) => ({ fileName: f.FileName, serverRelativeUrl: f.ServerRelativeUrl }))
      : []
  })), []);

  const loadData = React.useCallback(async () => {
    setLoading(true);
    try {
      const items = await service.getAll();
      setRawItems(items || []);
      setRecords(mapRecords(items || []));
      setError(null);
    } catch (err) {
      console.error('[CSAT] Failed to load data', err);
      setError('We could not load the latest CSAT entries. Please try again later.');
    } finally {
      setLoading(false);
    }
  }, [mapRecords, service]);

  React.useEffect(() => {
    void loadData();
  }, [loadData]);

  const sortedRecords = React.useMemo(() => {
    if (!isDateSorted) {
      return records;
    }
    return [...records].sort((a, b) => {
      const aTime = a.csatAquiredDate ? a.csatAquiredDate.getTime() : 0;
      const bTime = b.csatAquiredDate ? b.csatAquiredDate.getTime() : 0;
      const compareResult = aTime - bTime;
      return isDateSortedDescending ? -compareResult : compareResult;
    });
  }, [records, isDateSorted, isDateSortedDescending]);

  const totalPages = Math.max(1, Math.ceil(sortedRecords.length / PAGE_SIZE));

  const pagedRecords = React.useMemo(() => {
    const start = (currentPage - 1) * PAGE_SIZE;
    return sortedRecords.slice(start, start + PAGE_SIZE);
  }, [sortedRecords, currentPage]);

  React.useEffect(() => {
    setCurrentPage(1);
  }, [records, isDateSorted, isDateSortedDescending]);

  React.useEffect(() => {
    setCurrentPage(prev => Math.min(prev, totalPages));
  }, [totalPages]);

  const onDateColumnClick = React.useCallback(() => {
    setIsDateSortedDescending(isDateSorted ? !isDateSortedDescending : false);
    setIsDateSorted(true);
  }, [isDateSorted, isDateSortedDescending]);

  React.useEffect(() => {
    const checkPermissions = async (): Promise<void> => {
      try {
        const hasEditPermission = await ppoService.canUserEdit();
        setCanEdit(hasEditPermission);
      } catch (err) {
        console.error('[CSAT] Failed to check permissions', err);
        setCanEdit(false);
      }
    };
    void checkPermissions();
  }, [ppoService]);

  const formatDate = React.useCallback((date?: Date): string => {
    if (date instanceof Date && !isNaN(date.getTime())) {
      return date.toLocaleDateString();
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
    { key: 'colCSATValue', name: 'CSI', fieldName: 'csatValue', minWidth: 120, maxWidth: 120 },
    {
      key: 'colDate',
      name: 'Date',
      fieldName: 'csatAquiredDate',
      minWidth: 120,
      maxWidth: 120,
      isSorted: isDateSorted,
      isSortedDescending: isDateSortedDescending,
      onColumnClick: onDateColumnClick,
      onRender: (item: ICSATRecord) => formatDate(item.csatAquiredDate)
    },
    { key: 'colGoal', name: 'Goal', fieldName: 'goal', minWidth: 100, maxWidth: 100 },
    { key: 'colUSL', name: 'USL', fieldName: 'usl', minWidth: 80, maxWidth: 80 },
    { key: 'colLSL', name: 'LSL', fieldName: 'lsl', minWidth: 80, maxWidth: 80 },
    { key: 'colRemarks', name: 'Remarks', fieldName: 'remarks', minWidth: 250, maxWidth: 250 },
    {
      key: 'colAttachment',
      name: 'Attachment',
      fieldName: 'attachments',
      minWidth: 150,
      maxWidth: 200,
      onRender: (item: ICSATRecord) => (
        item.attachments && item.attachments.length > 0 ? (
          <div className={styles.attachmentCell}>
            {item.attachments.map(att => (
              <a
                key={att.fileName}
                href={att.serverRelativeUrl}
                target="_blank"
                rel="noopener noreferrer"
                title={att.fileName}
                className={styles.attachmentLink}
              >
                {att.fileName}
              </a>
            ))}
          </div>
        ) : (
          <span>—</span>
        )
      )
    }
  ], [formatDate, openEdit, canEdit, rawItems, isDateSorted, isDateSortedDescending, onDateColumnClick]);

  const tabs = React.useMemo(() => [
    { key: 'csat', label: 'Customer Satisfaction' }, 
    { key: 'workingDays', label: 'Working Days' }
  ], []);

  return (
    <section className={styles.tabShell} aria-labelledby="csat-heading">
      <Header
        title="CSAT & Working Days"
        iconName="Bullseye"
      />

      <div className={styles.tabButtons}>
        {tabs.map(tab => (
          <button
            key={tab.key}
            type="button"
            className={`${styles.tabButton} ${tab.key === 'csat' ? styles.activeTab : ''}`.trim()}
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
            <Spinner label="Loading CSAT data" />
          </div>
        ) : (
          <>
            <div className={styles.tableWrapper}>
              {sortedRecords.length > 0 ? (
                <DetailsList
                  items={pagedRecords}
                  columns={columns}
                  setKey="csat"
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  selectionMode={SelectionMode.none}
                />
              ) : (
                <div className={styles.emptyState}>No CSAT entries available.</div>
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
            title: editingItem ? 'Edit CSAT' : 'New CSAT'
          }}
          minWidth={500}
          maxWidth={600}
        >
          <NewEditCSAT
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

export default CSAT;
