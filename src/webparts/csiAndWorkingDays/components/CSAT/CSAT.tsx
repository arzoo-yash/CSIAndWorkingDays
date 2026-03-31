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
  PrimaryButton
} from '@fluentui/react';
import styles from './CSAT.module.scss';
import headerStyles from '../Header/Header.module.scss';
import type { ICSATProps, ICSATRecord } from './ICSAT';
import CSATService from '../../services/CSATService';
import { PPOApproverService } from '../../services/PPOApproverService';
import Header from '../Header';
import type { DashboardTabKey } from '../types';
import NewEditCSAT from './CSATForm';

const CSAT: React.FC<ICSATProps> = ({ context, onSelectTab }) => {
  const [records, setRecords] = React.useState<ICSATRecord[]>([]);
  const [rawItems, setRawItems] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [editingItem, setEditingItem] = React.useState<any | undefined>(undefined);
  const [canEdit, setCanEdit] = React.useState<boolean>(false);

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
    modified: item?.Modified ? new Date(item.Modified) : undefined
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
    { key: 'colCSATValue', name: 'CSAT Value', fieldName: 'csatValue', minWidth: 120, maxWidth: 120 },
    { key: 'colDate', name: 'Date', fieldName: 'csatAquiredDate', minWidth: 120, maxWidth: 120, onRender: (item: ICSATRecord) => formatDate(item.csatAquiredDate) },
    { key: 'colGoal', name: 'Goal', fieldName: 'goal', minWidth: 100, maxWidth: 100 },
    { key: 'colUSL', name: 'USL', fieldName: 'usl', minWidth: 80, maxWidth: 80 },
    { key: 'colLSL', name: 'LSL', fieldName: 'lsl', minWidth: 80, maxWidth: 80 },
    { key: 'colRemarks', name: 'Remarks', fieldName: 'remarks', minWidth: 250, maxWidth: 250 }
  ], [formatDate, openEdit, canEdit, rawItems]);

  const tabs = React.useMemo(() => [
    { key: 'csat', label: 'Customer Satisfaction' }, 
    { key: 'workingDays', label: 'No.Working Days' }
  ], []);

  return (
    <section className={styles.tabShell} aria-labelledby="csat-heading">
      <Header
        title="Admin Lists"
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
              {records.length > 0 ? (
                <DetailsList
                  items={records}
                  columns={columns}
                  setKey="csat"
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  selectionMode={SelectionMode.none}
                />
              ) : (
                <div className={styles.emptyState}>No CSAT entries available.</div>
              )}
            </div>
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
