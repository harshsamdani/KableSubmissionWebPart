import * as React from 'react';
import {
  Stack,
  DefaultButton,
  Text,
} from 'office-ui-fabric-react';
import { IContentItem, SectionType } from './IKableModels';
import ContentItemCard from './ContentItemCard';

export interface IContentSectionProps {
  section: SectionType;
  items: IContentItem[];
  onItemChange: (clientId: string, updated: IContentItem) => void;
  onItemDelete: (clientId: string) => void;
  onItemAdd: (section: SectionType) => void;
}

const SECTION_COLOURS: Record<SectionType, string> = {
  News: '#0078d4',
  Project: '#107c10',
  People: '#8764b8',
};

const SECTION_ICONS: Record<SectionType, string> = {
  News: '📰',
  Project: '📁',
  People: '👥',
};

export const ContentSection: React.FC<IContentSectionProps> = ({
  section,
  items,
  onItemChange,
  onItemDelete,
  onItemAdd,
}) => {
  const colour = SECTION_COLOURS[section];

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      {/* Section header */}
      <Stack
        horizontal
        verticalAlign="center"
        tokens={{ childrenGap: 8 }}
        styles={{
          root: {
            borderLeft: `4px solid ${colour}`,
            paddingLeft: 12,
            paddingTop: 4,
            paddingBottom: 4,
          }
        }}
      >
        <Text variant="xLarge" styles={{ root: { fontWeight: 600, color: colour } }}>
          {SECTION_ICONS[section]} {section}
        </Text>
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          ({items.length} item{items.length !== 1 ? 's' : ''})
        </Text>
      </Stack>

      {/* Items */}
      {items.length === 0 && (
        <Text variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic', paddingLeft: 16 } }}>
          No {section.toLowerCase()} items yet. Click &quot;Add {section} Item&quot; below.
        </Text>
      )}

      {items.map((item) => (
        <ContentItemCard
          key={item.clientId}
          item={item}
          onChange={(updated) => onItemChange(item.clientId, updated)}
          onDelete={() => onItemDelete(item.clientId)}
        />
      ))}

      {/* Add item button */}
      <Stack.Item>
        <DefaultButton
          iconProps={{ iconName: 'Add' }}
          text={`Add ${section} Item`}
          onClick={() => onItemAdd(section)}
          styles={{
            root: {
              borderColor: colour,
              color: colour,
            },
            rootHovered: {
              borderColor: colour,
              color: colour,
              background: `${colour}10`,
            }
          }}
        />
      </Stack.Item>
    </Stack>
  );
};

export default ContentSection;
