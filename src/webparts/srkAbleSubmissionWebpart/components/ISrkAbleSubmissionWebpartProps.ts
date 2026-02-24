import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISrkAbleSubmissionWebpartProps {
  spfxContext: WebPartContext;
  siteUrl: string;
  submissionsListName: string;
  contentListName: string;
}
