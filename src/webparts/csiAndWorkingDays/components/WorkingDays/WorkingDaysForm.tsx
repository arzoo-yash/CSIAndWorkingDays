import * as React from 'react';
import { Stack, TextField, PrimaryButton, DefaultButton, MessageBar, MessageBarType, Dropdown, IDropdownOption } from '@fluentui/react';
import type { IWorkingDaysFormProps, IWorkingDaysFormData } from './IWorkingDays';
import WorkingDaysService from '../../services/WorkingDaysService';

const WorkingDaysForm: React.FC<IWorkingDaysFormProps> = ({ context, item, isEdit, onSaved, onCancel }) => {
  const service = React.useMemo(() => new WorkingDaysService(context), [context]);

  const [form, setForm] = React.useState<IWorkingDaysFormData>(() => ({
    Title: item?.Title ?? item?.Month ?? undefined,
    Year: item?.Year ?? undefined,
    Days: item?.Days ?? undefined,
    ResourceAllocation: item?.ResourceAllocation ?? undefined
  }));

  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  const updateField = (key: keyof IWorkingDaysFormData, value: any) => setForm(prev => ({ ...prev, [key]: value }));

  const monthOptions: IDropdownOption[] = React.useMemo(() => [
    { key: 'January', text: 'January' },
    { key: 'February', text: 'February' },
    { key: 'March', text: 'March' },
    { key: 'April', text: 'April' },
    { key: 'May', text: 'May' },
    { key: 'June', text: 'June' },
    { key: 'July', text: 'July' },
    { key: 'August', text: 'August' },
    { key: 'September', text: 'September' },
    { key: 'October', text: 'October' },
    { key: 'November', text: 'November' },
    { key: 'December', text: 'December' }
  ], []);

  const yearOptions: IDropdownOption[] = React.useMemo(() => {
    const currentYear = new Date().getFullYear();
    const startYear = 2025;
    const endYear = currentYear + 10;
    const years: IDropdownOption[] = [];
    
    for (let year = startYear; year <= endYear; year++) {
      years.push({ key: year, text: year.toString() });
    }
    
    return years;
  }, []);

  const handleSave = async () => {
    setLoading(true);
    setError(null);
    try {
      const payload: any = { ...form };

      if (isEdit && item && item.Id) {
        const res = await service.update(item.Id, payload);
        onSaved?.(res);
      } else {
        const res = await service.create(payload);
        onSaved?.(res);
      }
    } catch (err) {
      console.error('[WorkingDaysForm] Save failed', err);
      setError('Failed to save working days entry. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: '24px', maxWidth: '500px', margin: '0 auto' }}>
      
      {error && (
        <MessageBar 
          messageBarType={MessageBarType.error} 
          styles={{ root: { marginBottom: '16px' } }}
        >
          {error}
        </MessageBar>
      )}
      
      <Stack tokens={{ childrenGap: 16 }}>
        <Dropdown
          label="Month"
          placeholder="Select a month"
          options={monthOptions}
          selectedKey={form.Title}
          onChange={(_, option) => updateField('Title', option?.key as string)}
          required
          styles={{ dropdown: { width: '100%' } }}
        />
        
        <Dropdown
          label="Year"
          placeholder="Select a year"
          options={yearOptions}
          selectedKey={form.Year}
          onChange={(_, option) => updateField('Year', option?.key as number)}
          required
          styles={{ dropdown: { width: '100%' } }}
        />
        
        <TextField
          label="Total Working Days"
          type="number"
          placeholder="Enter number of working days"
          value={form.Days !== undefined && form.Days !== null ? String(form.Days) : ''}
          onChange={(_, v) => updateField('Days', v === '' ? undefined : Number(v))}
          required
          min={1}
          max={31}
          styles={{ root: { width: '100%' } }}
        />
        
        <TextField
          label="Resource Allocation"
          type="number"
          placeholder="Enter resource allocation"
          value={form.ResourceAllocation !== undefined && form.ResourceAllocation !== null ? String(form.ResourceAllocation) : ''}
          onChange={(_, v) => updateField('ResourceAllocation', (v === '' || v === undefined) ? undefined : parseFloat(v))}
          required
          step={0.1}
          styles={{ root: { width: '100%' } }}
        />

        <Stack 
          horizontal 
          tokens={{ childrenGap: 12 }} 
          styles={{ root: { marginTop: '8px', justifyContent: 'flex-end' } }}
        >
          <DefaultButton 
            onClick={onCancel} 
            text="Cancel" 
            disabled={loading}
            styles={{ root: { minWidth: '100px' } }}
          />
          <PrimaryButton 
            onClick={handleSave} 
            text={isEdit ? 'Update' : 'Save'} 
            disabled={loading || !form.Title || !form.Year || !form.Days || form.ResourceAllocation === undefined || form.ResourceAllocation === null}
            styles={{ root: { minWidth: '100px' } }}
          />
        </Stack>
      </Stack>
    </div>
  );
};

export default WorkingDaysForm;
