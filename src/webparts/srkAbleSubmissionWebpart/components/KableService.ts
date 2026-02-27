import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/files';
import '@pnp/sp/files/folder';
import '@pnp/sp/folders';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IKableSubmissionForm } from './IKableModels';

export class KableService {
  private _sp: ReturnType<typeof spfi>;
  private _siteUrl: string;
  private _submissionsListName: string;
  private _contentListName: string;

  constructor(
    context: WebPartContext,
    siteUrl: string,
    submissionsListName: string,
    contentListName: string
  ) {
    this._siteUrl = siteUrl;
    this._submissionsListName = submissionsListName;
    this._contentListName = contentListName;

    this._sp = spfi(siteUrl).using(SPFx(context));
  }

  public async getGroupNameChoices(): Promise<string[]> {
    try {
      const field = await this._sp.web.lists
        .getByTitle(this._submissionsListName)
        .fields.getByInternalNameOrTitle('GroupName')();
      return (field as any).Choices || [];
    } catch (e) {
      console.error('Failed to load GroupName choices', e);
      return [];
    }
  }

  public async submitForm(form: IKableSubmissionForm): Promise<number> {
    // Step 1: Create the parent Kable Submissions item
    const submissionPayload: Record<string, any> = {
      Title: form.title,
      GroupName: form.groupName,
    };
    if (form.publishedDate) {
      submissionPayload['PublishedDate'] = form.publishedDate.toISOString();
    }

    const submissionResult = await this._sp.web.lists
      .getByTitle(this._submissionsListName)
      .items.add(submissionPayload);

    const submissionId: number = submissionResult.ID || submissionResult.data?.ID;

    // Step 2: Create each Kable Content item sequentially
    for (const item of form.items) {
      const contentPayload: Record<string, any> = {
        Title: '',
        KableSubmissionId: submissionId,  // lookup field — use Id suffix
        KableSection: item.section,
        Info: item.info,
        KableLayout: item.layout,
        SortOrder: item.sortOrder,
      };

      const contentResult = await this._sp.web.lists
        .getByTitle(this._contentListName)
        .items.add(contentPayload);

      const newItemId: number = contentResult.ID || contentResult.data?.ID;

      // Step 3: Upload image to SiteAssets document library and update KableImage field
      if (item.imageFile) {
        const fileName = item.imageFile.name;
        const fileBuffer = await this._readFileAsArrayBuffer(item.imageFile);
        const siteRelPath = new URL(this._siteUrl).pathname; // e.g. /sites/srkable
        const serverUrl = `https://${new URL(this._siteUrl).host}`;

        // Upload to SiteAssets with a unique prefix so filenames never collide.
        // Thumbnail fields require a document library URL — list attachment URLs do not render.
        const uniqueFileName = `kable-${newItemId}-${fileName}`;
        await this._sp.web
          .getFolderByServerRelativePath(`${siteRelPath}/SiteAssets`)
          .files.addUsingPath(uniqueFileName, fileBuffer, { Overwrite: true });

        const kableImageJson = JSON.stringify({
          type: 1,
          fileName: fileName,
          serverUrl: serverUrl,
          serverRelativeUrl: `${siteRelPath}/SiteAssets/${uniqueFileName}`,
        });

        await this._sp.web.lists
          .getByTitle(this._contentListName)
          .items.getById(newItemId)
          .update({ KableImage: kableImageJson });
      }
    }

    return submissionId;
  }

  private _readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }
}
