import * as React from 'react';
import {
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Text,
  Separator,
  DatePicker,
  DayOfWeek,
  IDatePickerStrings,
} from 'office-ui-fabric-react';
import { ISrkAbleSubmissionWebpartProps } from './ISrkAbleSubmissionWebpartProps';
import { IContentItem, IKableSubmissionForm, SectionType } from './IKableModels';
import { KableService } from './KableService';
import ContentSection from './ContentSection';

const DATE_PICKER_STRINGS: IDatePickerStrings = {
  months: ['January','February','March','April','May','June','July','August','September','October','November','December'],
  shortMonths: ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
  days: ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'],
  shortDays: ['S','M','T','W','T','F','S'],
  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker',
};

const SECTIONS: SectionType[] = ['News', 'Project', 'People'];

function makeEmptyItem(section: SectionType, order: number): IContentItem {
  return {
    clientId: `${section}-${Date.now()}-${Math.random()}`,
    section,
    info: '',
    imageFile: null,
    imagePreviewUrl: '',
    layout: 'Image Left, Info Right',
    sortOrder: order,
  };
}

function makeEmptyForm(): IKableSubmissionForm {
  return {
    title: '',
    groupName: '',
    publishedDate: null,
    items: [],
  };
}

type FormStatus = 'idle' | 'submitting' | 'success' | 'error';

const SrkAbleSubmissionWebpart: React.FC<ISrkAbleSubmissionWebpartProps> = (props) => {
  const { spfxContext, siteUrl, submissionsListName, contentListName } = props;

  const [form, setForm] = React.useState<IKableSubmissionForm>(makeEmptyForm);
  const [groupNameOptions, setGroupNameOptions] = React.useState<IDropdownOption[]>([]);
  const [loadingChoices, setLoadingChoices] = React.useState<boolean>(true);
  const [status, setStatus] = React.useState<FormStatus>('idle');
  const [errorMessage, setErrorMessage] = React.useState<string>('');
  const [submittedId, setSubmittedId] = React.useState<number | null>(null);
  const [validationErrors, setValidationErrors] = React.useState<Record<string, string>>({});

  const serviceRef = React.useRef<KableService | null>(null);

  // Initialise KableService and load GroupName choices on mount / when props change
  React.useEffect(() => {
    serviceRef.current = new KableService(spfxContext, siteUrl, submissionsListName, contentListName);
    setLoadingChoices(true);
    serviceRef.current.getGroupNameChoices().then((choices) => {
      setGroupNameOptions(choices.map((c) => ({ key: c, text: c })));
      setLoadingChoices(false);
    }).catch(() => {
      setLoadingChoices(false);
    });
  }, [siteUrl, submissionsListName, contentListName]);

  // ── Form field helpers ──────────────────────────────────────────────────────

  const setField = <K extends keyof IKableSubmissionForm>(key: K, value: IKableSubmissionForm[K]): void => {
    setForm((prev) => ({ ...prev, [key]: value }));
    // Clear validation error for this field
    if (validationErrors[key]) {
      setValidationErrors((prev) => { const next = { ...prev }; delete next[key]; return next; });
    }
  };

  const handleItemChange = (clientId: string, updated: IContentItem): void => {
    setForm((prev) => ({
      ...prev,
      items: prev.items.map((it) => it.clientId === clientId ? updated : it),
    }));
  };

  const handleItemDelete = (clientId: string): void => {
    setForm((prev) => ({
      ...prev,
      items: prev.items.filter((it) => it.clientId !== clientId),
    }));
  };

  const handleItemAdd = (section: SectionType): void => {
    const sectionItems = form.items.filter((it) => it.section === section);
    const nextOrder = sectionItems.length + 1;
    setForm((prev) => ({
      ...prev,
      items: [...prev.items, makeEmptyItem(section, nextOrder)],
    }));
  };

  // ── Validation ──────────────────────────────────────────────────────────────

  const validate = (): boolean => {
    const errors: Record<string, string> = {};
    if (!form.title.trim()) errors['title'] = 'Submission title is required.';
    if (!form.groupName) errors['groupName'] = 'Group name is required.';
    if (form.items.length === 0) errors['items'] = 'Add at least one content item.';
    setValidationErrors(errors);
    return Object.keys(errors).length === 0;
  };

  // ── Submit ──────────────────────────────────────────────────────────────────

  const handleSubmit = async (): Promise<void> => {
    if (!validate()) return;
    if (!serviceRef.current) return;

    setStatus('submitting');
    setErrorMessage('');

    try {
      const id = await serviceRef.current.submitForm(form);
      setSubmittedId(id);
      setStatus('success');
    } catch (err: any) {
      setErrorMessage(err?.message || 'An unexpected error occurred.');
      setStatus('error');
    }
  };

  const handleReset = (): void => {
    setForm(makeEmptyForm());
    setValidationErrors({});
    setStatus('idle');
    setSubmittedId(null);
    setErrorMessage('');
  };

  // ── Render ──────────────────────────────────────────────────────────────────

  if (status === 'success') {
    return (
      <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: 24, maxWidth: 800 } }}>
        <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
          Submission #{submittedId} created successfully! The page can now be published from the Kable Submissions list.
        </MessageBar>
        <DefaultButton text="Submit Another" iconProps={{ iconName: 'Add' }} onClick={handleReset} />
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 24 }} styles={{ root: { padding: 24, maxWidth: 800 } }}>

      {/* Header */}
      <Stack.Item>
        <Stack tokens={{ childrenGap: 6 }}>
          <Text variant="xxLarge" styles={{ root: { fontWeight: 700, color: '#323130' } }}>
            SRKable Newsletter Submission
          </Text>
          <Text variant="large" styles={{ root: { color: '#605e5c' } }}>
            Fill in the submission details below, then add content items for each section.
          </Text>
        </Stack>
      </Stack.Item>

      {/* Error banner */}
      {status === 'error' && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setStatus('idle')}>
          {errorMessage}
        </MessageBar>
      )}

      {/* Validation summary for items */}
      {validationErrors['items'] && (
        <MessageBar messageBarType={MessageBarType.warning}>
          {validationErrors['items']}
        </MessageBar>
      )}

      {/* ── Submission Header ── */}
      <Stack tokens={{ childrenGap: 16 }}
        styles={{ root: { background: '#f3f2f1', borderRadius: 4, padding: 20 } }}>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>Submission Details</Text>

        <TextField
          label="Submission Title"
          required
          placeholder="e.g. March 2025 Newsletter"
          value={form.title}
          onChange={(_, val) => setField('title', val || '')}
          errorMessage={validationErrors['title']}
        />

        <Stack horizontal tokens={{ childrenGap: 24 }} wrap>
          <Stack.Item grow={2} styles={{ root: { minWidth: 200 } }}>
            {loadingChoices ? (
              <Spinner size={SpinnerSize.small} label="Loading groups..." />
            ) : (
              <Dropdown
                label="Group Name"
                required
                placeholder="Select a group"
                options={groupNameOptions}
                selectedKey={form.groupName || null}
                onChange={(_, opt) => opt && setField('groupName', opt.key as string)}
                errorMessage={validationErrors['groupName']}
              />
            )}
          </Stack.Item>

          <Stack.Item grow={1} styles={{ root: { minWidth: 180 } }}>
            <DatePicker
              label="Published Date"
              placeholder="Select a date (optional)"
              firstDayOfWeek={DayOfWeek.Monday}
              strings={DATE_PICKER_STRINGS}
              value={form.publishedDate || undefined}
              onSelectDate={(date) => setField('publishedDate', date || null)}
              formatDate={(d) => d ? d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }) : ''}
            />
          </Stack.Item>
        </Stack>
      </Stack>

      <Separator />

      {/* ── Content Sections ── */}
      {SECTIONS.map((section) => (
        <ContentSection
          key={section}
          section={section}
          items={form.items.filter((it) => it.section === section)}
          onItemChange={handleItemChange}
          onItemDelete={handleItemDelete}
          onItemAdd={handleItemAdd}
        />
      ))}

      <Separator />

      {/* ── Submit ── */}
      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
        <PrimaryButton
          text={status === 'submitting' ? 'Submitting...' : 'Submit'}
          iconProps={{ iconName: status === 'submitting' ? undefined : 'Send' }}
          onClick={handleSubmit}
          disabled={status === 'submitting'}
          styles={{ root: { minWidth: 120, height: 40 } }}
        />
        {status === 'submitting' && <Spinner size={SpinnerSize.small} />}
        <DefaultButton
          text="Clear Form"
          onClick={handleReset}
          disabled={status === 'submitting'}
          styles={{ root: { minWidth: 120, height: 40 } }}
        />
      </Stack>

    </Stack>
  );
};

export default SrkAbleSubmissionWebpart;
