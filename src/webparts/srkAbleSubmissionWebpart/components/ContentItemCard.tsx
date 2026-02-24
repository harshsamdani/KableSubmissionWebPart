import * as React from 'react';
import {
  Stack,
  Dropdown,
  IDropdownOption,
  SpinButton,
  IconButton,
  Label,
  Text,
} from 'office-ui-fabric-react';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { IContentItem, LAYOUT_OPTIONS } from './IKableModels';

export interface IContentItemCardProps {
  item: IContentItem;
  onChange: (updated: IContentItem) => void;
  onDelete: () => void;
}

const layoutDropdownOptions: IDropdownOption[] = LAYOUT_OPTIONS.map(l => ({ key: l, text: l }));

export const ContentItemCard: React.FC<IContentItemCardProps> = ({ item, onChange, onDelete }) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleImageChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const previewUrl = URL.createObjectURL(file);
    onChange({ ...item, imageFile: file, imagePreviewUrl: previewUrl });
  };

  const handleRemoveImage = (): void => {
    if (item.imagePreviewUrl) URL.revokeObjectURL(item.imagePreviewUrl);
    onChange({ ...item, imageFile: null, imagePreviewUrl: '' });
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  return (
    <Stack
      tokens={{ childrenGap: 12 }}
      styles={{
        root: {
          border: '1px solid #e1e1e1',
          borderRadius: 4,
          padding: 16,
          background: '#fafafa',
          position: 'relative',
        }
      }}
    >
      {/* Delete button */}
      <IconButton
        iconProps={{ iconName: 'Cancel' }}
        title="Remove item"
        ariaLabel="Remove item"
        onClick={onDelete}
        styles={{
          root: {
            position: 'absolute',
            top: 8,
            right: 8,
            color: '#a4262c',
          }
        }}
      />

      {/* Rich text — Info */}
      <Stack.Item>
        <Label required>Content / Info</Label>
        <div style={{ minHeight: 150, border: '1px solid #c8c6c4', borderRadius: 2 }}>
          <RichText
            value={item.info}
            onChange={(newVal: string) => {
              onChange({ ...item, info: newVal || '' });
              return newVal;
            }}
            isEditMode={true}
          />
        </div>
      </Stack.Item>

      {/* Image upload */}
      <Stack.Item>
        <Label>Image</Label>
        <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            style={{ display: 'none' }}
            onChange={handleImageChange}
          />
          <IconButton
            iconProps={{ iconName: 'Photo2' }}
            text="Choose image"
            onClick={() => fileInputRef.current && fileInputRef.current.click()}
            styles={{ root: { border: '1px solid #c8c6c4', borderRadius: 2, padding: '4px 12px', height: 32 } }}
          />
          {item.imagePreviewUrl && (
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <img
                src={item.imagePreviewUrl}
                alt="preview"
                style={{ height: 48, width: 'auto', borderRadius: 2, objectFit: 'cover' }}
              />
              <Text variant="small">{item.imageFile ? item.imageFile.name : ''}</Text>
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Remove image"
                onClick={handleRemoveImage}
                styles={{ root: { color: '#a4262c' } }}
              />
            </Stack>
          )}
        </Stack>
      </Stack.Item>

      {/* Layout + Sort Order in a row */}
      <Stack horizontal tokens={{ childrenGap: 24 }} wrap>
        <Stack.Item grow={3}>
          <Dropdown
            label="Layout"
            required
            options={layoutDropdownOptions}
            selectedKey={item.layout}
            onChange={(_, opt) => opt && onChange({ ...item, layout: opt.key as string })}
          />
        </Stack.Item>
        <Stack.Item grow={1} styles={{ root: { minWidth: 100 } }}>
          <Label>Sort Order</Label>
          <SpinButton
            value={String(item.sortOrder)}
            min={1}
            max={999}
            step={1}
            onValidate={(val) => {
              const n = parseInt(val, 10);
              const clamped = isNaN(n) ? 1 : Math.max(1, n);
              onChange({ ...item, sortOrder: clamped });
              return String(clamped);
            }}
            onIncrement={(val) => {
              const n = parseInt(val, 10) + 1;
              onChange({ ...item, sortOrder: n });
              return String(n);
            }}
            onDecrement={(val) => {
              const n = Math.max(1, parseInt(val, 10) - 1);
              onChange({ ...item, sortOrder: n });
              return String(n);
            }}
          />
        </Stack.Item>
      </Stack>
    </Stack>
  );
};

export default ContentItemCard;
