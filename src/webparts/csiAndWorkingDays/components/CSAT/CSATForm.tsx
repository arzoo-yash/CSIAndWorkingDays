import * as React from 'react';
import { Stack, TextField, PrimaryButton, DefaultButton, MessageBar, MessageBarType, DatePicker, DayOfWeek, IconButton, Label } from '@fluentui/react';
import type { INewEditCSATProps, INewEditCSATFormData } from './ICSATForm';
import CSATService from '../../services/CSATService';
import ProjectMetricsService from '../../services/ProjectMetricsService';

interface IExistingAttachment {
  fileName: string;
  serverRelativeUrl: string;
}

const NewEditCSAT: React.FC<INewEditCSATProps> = ({ context, item, isEdit, onSaved, onCancel }) => {
  const service = React.useMemo(() => new CSATService(context), [context]);
  const metricsService = React.useMemo(() => new ProjectMetricsService(context), [context]);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const [form, setForm] = React.useState<INewEditCSATFormData>(() => ({
    Title: item?.Title ?? undefined,
    CSATAquiredDate: item?.CSATAquiredDate ? new Date(item.CSATAquiredDate) : undefined,
    Goal: item?.Goal ?? undefined,
    USL: item?.USL ?? undefined,
    LSL: item?.LSL ?? undefined,
    Remarks: item?.Remarks ?? undefined
  }));

  const [existingAttachments, setExistingAttachments] = React.useState<IExistingAttachment[]>(() =>
    Array.isArray(item?.AttachmentFiles)
      ? item.AttachmentFiles.map((f: any) => ({ fileName: f.FileName, serverRelativeUrl: f.ServerRelativeUrl }))
      : []
  );
  const [attachmentsToRemove, setAttachmentsToRemove] = React.useState<string[]>([]);
  const [newAttachments, setNewAttachments] = React.useState<File[]>([]);

  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [validationErrors, setValidationErrors] = React.useState<{ [key: string]: string }>({});
  const [ppoValidation, setPpoValidation] = React.useState<{ isValid: boolean; message?: string }>({ isValid: true });
  const [checkingPPO, setCheckingPPO] = React.useState(false);

  // Validate PPO and fetch the latest CSAT metrics from ProjectMetrics on component mount
  React.useEffect(() => {
    const validateAndFetchMetrics = async () => {
      setCheckingPPO(true);
      setError(null); // Clear any previous errors
      try {
        const validation = await metricsService.validateAndGetCSATMetrics();
        setPpoValidation({ isValid: validation.isValid, message: validation.message });

        if (validation.isValid && validation.metrics) {
          // Update form with fetched metrics values (Goal, USL, LSL)
          setForm(prev => ({
            ...prev,
            Goal: validation.metrics?.Goal ?? prev.Goal,
            USL: validation.metrics?.USL ?? prev.USL,
            LSL: validation.metrics?.LSL ?? prev.LSL
          }));
        }
        // No need to set error here - ppoValidation.message will display it
      } catch (err) {
        console.error('[NewEditCSAT] Failed to validate PPO and fetch metrics', err);
        setPpoValidation({ 
          isValid: false, 
          message: 'Error checking PPO status. Please try again.' 
        });
        // No need to set error here - ppoValidation.message will display it
      } finally {
        setCheckingPPO(false);
      }
    };

    // Fetch metrics if:
    // 1. Creating new item (!isEdit)
    // 2. Editing but Goal/USL/LSL are missing from the item (to backfill from latest metrics)
    const hasExistingMetrics = item?.Goal || item?.USL || item?.LSL;
    const needsMetrics = !isEdit || (isEdit && !hasExistingMetrics);
    
    if (needsMetrics) {
      void validateAndFetchMetrics();
    } else {
      // When editing with existing metrics, consider PPO as valid
      setPpoValidation({ isValid: true });
    }
  }, [metricsService, isEdit, item]);

  const updateField = (key: keyof INewEditCSATFormData, value: any) => {
    setForm(prev => ({ ...prev, [key]: value }));
    // Clear validation error for this field when user updates it
    if (validationErrors[key]) {
      setValidationErrors(prev => {
        const updated = { ...prev };
        delete updated[key];
        return updated;
      });
    }
  };

  const validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    
    if (!form.Title || form.Title.trim() === '') {
      errors.Title = 'CSI is required';
    }
    
    if (!form.CSATAquiredDate) {
      errors.CSATAquiredDate = 'CSAT Acquired Date is required';
    }
    
    setValidationErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const handleFilesSelected = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      const selected: File[] = [];
      for (let i = 0; i < files.length; i++) {
        selected.push(files[i]);
      }
      setNewAttachments(prev => [...prev, ...selected]);
    }
    // Reset the input so the same file can be re-selected if removed and re-added
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const removeNewAttachment = (fileName: string) => {
    setNewAttachments(prev => prev.filter(f => f.name !== fileName));
  };

  const removeExistingAttachment = (fileName: string) => {
    setExistingAttachments(prev => prev.filter(a => a.fileName !== fileName));
    setAttachmentsToRemove(prev => [...prev, fileName]);
  };

  const saveAttachments = async (itemId: number): Promise<void> => {
    for (const fileName of attachmentsToRemove) {
      try {
        await service.deleteAttachment(itemId, fileName);
      } catch (err) {
        console.error(`[NewEditCSAT] Failed to remove attachment "${fileName}"`, err);
      }
    }

    for (const file of newAttachments) {
      try {
        const content = await file.arrayBuffer();
        await service.addAttachment(itemId, file.name, content);
      } catch (err) {
        console.error(`[NewEditCSAT] Failed to upload attachment "${file.name}"`, err);
      }
    }
  };

  const handleSave = async () => {
    // Check PPO validation before proceeding
    if (!ppoValidation.isValid) {
      setError(ppoValidation.message || 'PPO validation failed. Cannot save CSAT entry.');
      return;
    }

    if (!validateForm()) {
      setError('Please fill in all required fields.');
      return;
    }

    setLoading(true);
    setError(null);
    try {
      const payload: any = { ...form };
      if (payload.CSATAquiredDate instanceof Date) {
        payload.CSATAquiredDate = payload.CSATAquiredDate.toISOString();
      }

      if (isEdit && item && item.Id) {
        const res = await service.update(item.Id, payload);
        await saveAttachments(item.Id);
        onSaved?.(res);
      } else {
        const res = await service.create(payload);
        const newItemId = res?.Id || res?.data?.Id;
        if (newItemId) {
          await saveAttachments(newItemId);
        }
        onSaved?.(res);
      }
    } catch (err) {
      console.error('[NewEditCSAT] Save failed', err);
      setError('Failed to save CSAT entry. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: '20px' }}>
      {checkingPPO && (
        <MessageBar messageBarType={MessageBarType.info}>
          Validating Process Performance Objectives...
        </MessageBar>
      )}
      {!checkingPPO && !ppoValidation.isValid && ppoValidation.message && (
        <MessageBar messageBarType={MessageBarType.severeWarning}>
          {ppoValidation.message}
        </MessageBar>
      )}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField
          label="CSI"
          required
          value={form.Title ?? ''}
          onChange={(_, v) => updateField('Title', v)}
          errorMessage={validationErrors.Title}
          disabled={checkingPPO || !ppoValidation.isValid}
        />

        <DatePicker
          label="CSAT Acquired Date"
          isRequired
          firstDayOfWeek={DayOfWeek.Sunday}
          placeholder="Select a date"
          value={form.CSATAquiredDate ? (form.CSATAquiredDate instanceof Date ? form.CSATAquiredDate : new Date(String(form.CSATAquiredDate))) : undefined}
          onSelectDate={(d) => updateField('CSATAquiredDate', d)}
          disabled={checkingPPO || !ppoValidation.isValid}
        />
        {validationErrors.CSATAquiredDate && (
          <MessageBar messageBarType={MessageBarType.error}>{validationErrors.CSATAquiredDate}</MessageBar>
        )}

        <Stack horizontal tokens={{ childrenGap: 12 }}>
          <TextField
            label="Goal"
            value={form.Goal ?? ''}
            disabled
            readOnly
            styles={{ root: { flex: 1 } }}
          />
          <TextField
            label="USL"
            value={form.USL !== undefined ? String(form.USL) : ''}
            disabled
            readOnly
            styles={{ root: { flex: 1 } }}
          />
          <TextField
            label="LSL"
            value={form.LSL !== undefined ? String(form.LSL) : ''}
            disabled
            readOnly
            styles={{ root: { flex: 1 } }}
          />
        </Stack>

        <TextField
          label="Remarks"
          multiline
          rows={4}
          value={form.Remarks}
          onChange={(_, v) => updateField('Remarks', v)}
          disabled={checkingPPO || !ppoValidation.isValid}
        />

        <Stack tokens={{ childrenGap: 6 }}>
          <Label>Attachment</Label>
          {existingAttachments.map(att => (
            <Stack
              key={att.fileName}
              horizontal
              verticalAlign="center"
              tokens={{ childrenGap: 8 }}
            >
              <a href={att.serverRelativeUrl} target="_blank" rel="noopener noreferrer">{att.fileName}</a>
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Remove attachment"
                ariaLabel="Remove attachment"
                onClick={() => removeExistingAttachment(att.fileName)}
                disabled={checkingPPO || !ppoValidation.isValid}
              />
            </Stack>
          ))}
          {newAttachments.map(file => (
            <Stack
              key={file.name}
              horizontal
              verticalAlign="center"
              tokens={{ childrenGap: 8 }}
            >
              <span>{file.name}</span>
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Remove attachment"
                ariaLabel="Remove attachment"
                onClick={() => removeNewAttachment(file.name)}
                disabled={checkingPPO || !ppoValidation.isValid}
              />
            </Stack>
          ))}
          <input
            ref={fileInputRef}
            type="file"
            multiple
            style={{ display: 'none' }}
            onChange={handleFilesSelected}
          />
          <DefaultButton
            iconProps={{ iconName: 'Attach' }}
            text="Add Attachment"
            onClick={() => fileInputRef.current?.click()}
            disabled={checkingPPO || !ppoValidation.isValid}
            styles={{ root: { width: 'fit-content' } }}
          />
        </Stack>

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton onClick={onCancel} text="Cancel" disabled={loading || checkingPPO} />
          <PrimaryButton 
            onClick={handleSave} 
            text={isEdit ? 'Update' : 'Save'} 
            disabled={loading || checkingPPO || (!isEdit && !ppoValidation.isValid)} 
          />
        </Stack>
      </Stack>
    </div>
  );
};

export default NewEditCSAT;
