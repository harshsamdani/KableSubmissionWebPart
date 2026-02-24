export type SectionType = 'News' | 'Project' | 'People';

export const LAYOUT_OPTIONS = [
  'Image Left, Info Right',
  'Image Right, Info Left',
  'Image Top, Info Below',
  'Full Width Banner',
  'Info Only',
];

export interface IContentItem {
  clientId: string;
  section: SectionType;
  info: string;
  imageFile: File | null;
  imagePreviewUrl: string;
  layout: string;
  sortOrder: number;
}

export interface IKableSubmissionForm {
  title: string;
  groupName: string;
  publishedDate: Date | null;
  items: IContentItem[];
}
