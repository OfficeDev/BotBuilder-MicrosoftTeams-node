import { CardFactory, Attachment } from 'botbuilder';
import * as models from '../schema';

export class TeamsFactory extends CardFactory {
  public static contentTypes: any = {
    ...CardFactory.contentTypes,
    o365ConnectorCard: 'application/vnd.microsoft.teams.card.o365connector',
    fileConsentCard: 'application/vnd.microsoft.teams.card.file.consent',
    fileDownloadInfo: 'application/vnd.microsoft.teams.file.download.info',
    fileInfoCard: 'application/vnd.microsoft.teams.card.file.info'
  };

  public static isFileDownloadInfoAttachment (attachment: Attachment): attachment is models.FileDownloadInfoAttachment {
    return attachment.contentType === TeamsFactory.contentTypes.fileDownloadInfo;
  }

  public static o365ConnectorCard (content: models.O365ConnectorCard): Attachment {
    return {
      contentType: TeamsFactory.contentTypes.o365ConnectorCard,
      content
    };
  }

  public static fileConsentCard (fileName: string, content: models.FileConsentCard): Attachment {
    return {
      contentType: TeamsFactory.contentTypes.fileConsentCard,
      name: fileName,
      content
    };
  }

  public static fileInfoCard (fileName: string, contentUrl: string, content: models.FileInfoCard): Attachment {
    return {
      contentType: TeamsFactory.contentTypes.fileInfoCard,
      name: fileName,
      contentUrl,
      content
    };
  }
}
