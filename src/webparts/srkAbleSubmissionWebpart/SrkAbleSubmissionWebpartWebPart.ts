import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import SrkAbleSubmissionWebpart from './components/SrkAbleSubmissionWebpart';
import { ISrkAbleSubmissionWebpartProps } from './components/ISrkAbleSubmissionWebpartProps';

export interface ISrkAbleSubmissionWebpartWebPartProps {
  siteUrl: string;
  submissionsListName: string;
  contentListName: string;
}

export default class SrkAbleSubmissionWebpartWebPart extends BaseClientSideWebPart<ISrkAbleSubmissionWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISrkAbleSubmissionWebpartProps> = React.createElement(
      SrkAbleSubmissionWebpart,
      {
        spfxContext: this.context,
        siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
        submissionsListName: this.properties.submissionsListName || 'Kable Submissions',
        contentListName: this.properties.contentListName || 'Kable Content',
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the SRKable Submission Form'
          },
          groups: [
            {
              groupName: 'List Settings',
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL',
                  description: 'SharePoint site URL where the Kable lists are located. Leave blank to use the current site.',
                  placeholder: 'https://yourtenant.sharepoint.com/sites/yoursite'
                }),
                PropertyPaneTextField('submissionsListName', {
                  label: 'Submissions List Name',
                  description: 'Display name of the Kable Submissions (parent) list.',
                  placeholder: 'Kable Submissions'
                }),
                PropertyPaneTextField('contentListName', {
                  label: 'Content List Name',
                  description: 'Display name of the Kable Content (child) list.',
                  placeholder: 'Kable Content'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
