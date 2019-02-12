/*
 * Code generated by Microsoft (R) AutoRest Code Generator.
 * Changes may cause incorrect behavior and will be lost if the code is
 * regenerated.
 */

import * as teams from '../extension'
import * as builder from 'botbuilder'
import { ServiceClientOptions } from 'ms-rest-js';
import * as msRest from 'ms-rest-js';


/**
 * @interface
 * An interface representing ChannelInfo.
 * A channel info object which decribes the channel.
 *
 */
export interface ChannelInfo {
  /**
   * @member {string} [id] Unique identifier representing a channel
   */
  id?: string;
  /**
   * @member {string} [name] Name of the channel
   */
  name?: string;
}

/**
 * @interface
 * An interface representing ConversationList.
 * List of channels under a team
 *
 */
export interface ConversationList {
  /**
   * @member {ChannelInfo[]} [conversations]
   */
  conversations?: ChannelInfo[];
}

/**
 * @interface
 * An interface representing TeamDetails.
 * Details related to a team
 *
 */
export interface TeamDetails {
  /**
   * @member {string} [id] Unique identifier representing a team
   */
  id?: string;
  /**
   * @member {string} [name] Name of team.
   */
  name?: string;
  /**
   * @member {string} [aadGroupId] Azure Active Directory (AAD) Group Id for
   * the team.
   */
  aadGroupId?: string;
}

/**
 * @interface
 * An interface representing TeamInfo.
 * Describes a team
 *
 */
export interface TeamInfo {
  /**
   * @member {string} [id] Unique identifier representing a team
   */
  id?: string;
  /**
   * @member {string} [name] Name of team.
   */
  name?: string;
}

/**
 * @interface
 * An interface representing NotificationInfo.
 * Specifies if a notification is to be sent for the mentions.
 *
 */
export interface NotificationInfo {
  /**
   * @member {boolean} [alert] true if notification is to be sent to the user,
   * false otherwise.
   */
  alert?: boolean;
}

/**
 * @interface
 * An interface representing TenantInfo.
 * Describes a tenant
 *
 */
export interface TenantInfo {
  /**
   * @member {string} [id] Unique identifier representing a tenant
   */
  id?: string;
}

/**
 * @interface
 * An interface representing TeamsChannelData.
 * List of channels under a team
 *
 */
export interface TeamsChannelData {
  /**
   * @member {ChannelInfo} [channel]
   */
  channel?: ChannelInfo;
  /**
   * @member {string} [eventType] Type of event.
   */
  eventType?: string;
  /**
   * @member {TeamInfo} [team]
   */
  team?: TeamInfo;
  /**
   * @member {NotificationInfo} [notification]
   */
  notification?: NotificationInfo;
  /**
   * @member {TenantInfo} [tenant]
   */
  tenant?: TenantInfo;
}

/**
 * @interface
 * An interface representing TeamsChannelAccount.
 * Teams channel account detailing user Azure Active Directory details.
 *
 * @extends builder.ChannelAccount
 */
export interface TeamsChannelAccount extends builder.ChannelAccount {
  /**
   * @member {string} [aadObjectId] Unique Azure Active Directory object Id.
   */
  aadObjectId?: string;
  /**
   * @member {string} [givenName] Given name part of the user name.
   */
  givenName?: string;
  /**
   * @member {string} [surname] Surname part of the user name.
   */
  surname?: string;
  /**
   * @member {string} [email] Email Id of the user.
   */
  email?: string;
  /**
   * @member {string} [userPrincipalName] Unique user principal name
   */
  userPrincipalName?: string;
}

/**
 * @interface
 * An interface representing O365ConnectorCardFact.
 * O365 connector card fact
 *
 */
export interface O365ConnectorCardFact {
  /**
   * @member {string} [name] Display name of the fact
   */
  name?: string;
  /**
   * @member {string} [value] Display value for the fact
   */
  value?: string;
}

/**
 * @interface
 * An interface representing O365ConnectorCardImage.
 * O365 connector card image
 *
 */
export interface O365ConnectorCardImage {
  /**
   * @member {string} [image] URL for the image
   */
  image?: string;
  /**
   * @member {string} [title] Alternative text for the image
   */
  title?: string;
}

/**
 * @interface
 * An interface representing O365ConnectorCardSection.
 * O365 connector card section
 *
 */
export interface O365ConnectorCardSection {
  /**
   * @member {string} [title] Title of the section
   */
  title?: string;
  /**
   * @member {string} [text] Text for the section
   */
  text?: string;
  /**
   * @member {string} [activityTitle] Activity title
   */
  activityTitle?: string;
  /**
   * @member {string} [activitySubtitle] Activity subtitle
   */
  activitySubtitle?: string;
  /**
   * @member {string} [activityText] Activity text
   */
  activityText?: string;
  /**
   * @member {string} [activityImage] Activity image
   */
  activityImage?: string;
  /**
   * @member {ActivityImageType} [activityImageType] Describes how Activity
   * image is rendered. Possible values include: 'avatar', 'article'
   */
  activityImageType?: ActivityImageType;
  /**
   * @member {boolean} [markdown] Use markdown for all text contents. Default
   * vaule is true.
   */
  markdown?: boolean;
  /**
   * @member {O365ConnectorCardFact[]} [facts] Set of facts for the current
   * section
   */
  facts?: O365ConnectorCardFact[];
  /**
   * @member {O365ConnectorCardImage[]} [images] Set of images for the current
   * section
   */
  images?: O365ConnectorCardImage[];
  /**
   * @member {O365ConnectorCardActionBase[]} [potentialAction] Set of actions
   * for the current section
   */
  potentialAction?: teams.O365ConnectorCardActionBase[];
}

/**
 * @interface
 * An interface representing O365ConnectorCard.
 * O365 connector card
 *
 */
export interface O365ConnectorCard {
  /**
   * @member {string} [title] Title of the item
   */
  title?: string;
  /**
   * @member {string} [text] Text for the card
   */
  text?: string;
  /**
   * @member {string} [summary] Summary for the card
   */
  summary?: string;
  /**
   * @member {string} [themeColor] Theme color for the card
   */
  themeColor?: string;
  /**
   * @member {O365ConnectorCardSection[]} [sections] Set of sections for the
   * current card
   */
  sections?: O365ConnectorCardSection[];
  /**
   * @member {O365ConnectorCardActionBase[]} [potentialAction] Set of actions
   * for the current card
   */
  potentialAction?: teams.O365ConnectorCardActionBase[];
}

/**
 * @interface
 * An interface representing O365ConnectorCardViewAction.
 * O365 connector card ViewAction action
 *
 * @extends teams.O365ConnectorCardActionBase
 */
export interface O365ConnectorCardViewAction extends teams.O365ConnectorCardActionBase {
  /**
   * @member {string[]} [target] Target urls, only the first url effective for
   * card button
   */
  target?: string[];
}

/**
 * @interface
 * An interface representing O365ConnectorCardOpenUriTarget.
 * O365 connector card OpenUri target
 *
 */
export interface O365ConnectorCardOpenUriTarget {
  /**
   * @member {Os} [os] Target operating system. Possible values include:
   * 'default', 'iOS', 'android', 'windows'
   */
  os?: Os;
  /**
   * @member {string} [uri] Target url
   */
  uri?: string;
}

/**
 * @interface
 * An interface representing O365ConnectorCardOpenUri.
 * O365 connector card OpenUri action
 *
 * @extends teams.O365ConnectorCardActionBase
 */
export interface O365ConnectorCardOpenUri extends teams.O365ConnectorCardActionBase {
  /**
   * @member {O365ConnectorCardOpenUriTarget[]} [targets] Target os / urls
   */
  targets?: O365ConnectorCardOpenUriTarget[];
}

/**
 * @interface
 * An interface representing O365ConnectorCardHttpPOST.
 * O365 connector card HttpPOST action
 *
 * @extends teams.O365ConnectorCardActionBase
 */
export interface O365ConnectorCardHttpPOST extends teams.O365ConnectorCardActionBase {
  /**
   * @member {string} [body] Content to be posted back to bots via invoke
   */
  body?: string;
}

/**
 * @interface
 * An interface representing O365ConnectorCardActionCard.
 * O365 connector card ActionCard action
 *
 * @extends teams.O365ConnectorCardActionBase
 */
export interface O365ConnectorCardActionCard extends teams.O365ConnectorCardActionBase {
  /**
   * @member {O365ConnectorCardInputBase[]} [inputs] Set of inputs contained in
   * this ActionCard whose each item can be in any subtype of
   * teams.O365ConnectorCardInputBase
   */
  inputs?: teams.O365ConnectorCardInputBase[];
  /**
   * @member {O365ConnectorCardActionBase[]} [actions] Set of actions contained
   * in this ActionCard whose each item can be in any subtype of
   * teams.O365ConnectorCardActionBase except O365ConnectorCardActionCard, as nested
   * ActionCard is forbidden.
   */
  actions?: teams.O365ConnectorCardActionBase[];
}

/**
 * @interface
 * An interface representing O365ConnectorCardTextInput.
 * O365 connector card text input
 *
 * @extends teams.O365ConnectorCardInputBase
 */
export interface O365ConnectorCardTextInput extends teams.O365ConnectorCardInputBase {
  /**
   * @member {boolean} [isMultiline] Define if text input is allowed for
   * multiple lines. Default value is false.
   */
  isMultiline?: boolean;
  /**
   * @member {number} [maxLength] Maximum length of text input. Default value
   * is unlimited.
   */
  maxLength?: number;
}

/**
 * @interface
 * An interface representing O365ConnectorCardDateInput.
 * O365 connector card date input
 *
 * @extends teams.O365ConnectorCardInputBase
 */
export interface O365ConnectorCardDateInput extends teams.O365ConnectorCardInputBase {
  /**
   * @member {boolean} [includeTime] Include time input field. Default value
   * is false (date only).
   */
  includeTime?: boolean;
}

/**
 * @interface
 * An interface representing O365ConnectorCardMultichoiceInputChoice.
 * O365O365 connector card multiple choice input item
 *
 */
export interface O365ConnectorCardMultichoiceInputChoice {
  /**
   * @member {string} [display] The text rednered on ActionCard.
   */
  display?: string;
  /**
   * @member {string} [value] The value received as results.
   */
  value?: string;
}

/**
 * @interface
 * An interface representing O365ConnectorCardMultichoiceInput.
 * O365 connector card multiple choice input
 *
 * @extends teams.O365ConnectorCardInputBase
 */
export interface O365ConnectorCardMultichoiceInput extends teams.O365ConnectorCardInputBase {
  /**
   * @member {O365ConnectorCardMultichoiceInputChoice[]} [choices] Set of
   * choices whose each item can be in any subtype of
   * O365ConnectorCardMultichoiceInputChoice.
   */
  choices?: O365ConnectorCardMultichoiceInputChoice[];
  /**
   * @member {Style} [style] Choice item rendering style. Default valud is
   * 'compact'. Possible values include: 'compact', 'expanded'
   */
  style?: Style;
  /**
   * @member {boolean} [isMultiSelect] Define if this input field allows
   * multiple selections. Default value is false.
   */
  isMultiSelect?: boolean;
}

/**
 * @interface
 * An interface representing O365ConnectorCardActionQuery.
 * O365 connector card HttpPOST invoke query
 *
 */
export interface O365ConnectorCardActionQuery {
  /**
   * @member {string} [body] The results of body string defined in
   * IO365ConnectorCardHttpPOST with substituted input values
   */
  body?: string;
  /**
   * @member {string} [actionId] Action Id associated with the HttpPOST action
   * button triggered, defined in teams.O365ConnectorCardActionBase.
   */
  actionId?: string;
}

/**
 * @interface
 * An interface representing SigninStateVerificationQuery.
 * Signin state (part of signin action auth flow) verification invoke query
 *
 */
export interface SigninStateVerificationQuery {
  /**
   * @member {string} [state] The state string originally received when the
   * signin web flow is finished with a state posted back to client via tab SDK
   * microsoftTeams.authentication.notifySuccess(state)
   */
  state?: string;
}

/**
 * @interface
 * An interface representing MessagingExtensionQueryOptions.
 * Messaging extension query options
 *
 */
export interface MessagingExtensionQueryOptions {
  /**
   * @member {number} [skip] Number of entities to skip
   */
  skip?: number;
  /**
   * @member {number} [count] Number of entities to fetch
   */
  count?: number;
}

/**
 * @interface
 * An interface representing MessagingExtensionParameter.
 * Messaging extension query parameters
 *
 */
export interface MessagingExtensionParameter {
  /**
   * @member {string} [name] Name of the parameter
   */
  name?: string;
  /**
   * @member {any} [value] Value of the parameter
   */
  value?: any;
}

/**
 * @interface
 * An interface representing MessagingExtensionQuery.
 * Messaging extension query
 *
 */
export interface MessagingExtensionQuery {
  /**
   * @member {string} [commandId] Id of the command assigned by Bot
   */
  commandId?: string;
  /**
   * @member {MessagingExtensionParameter[]} [parameters] Parameters for the
   * query
   */
  parameters?: MessagingExtensionParameter[];
  /**
   * @member {MessagingExtensionQueryOptions} [queryOptions]
   */
  queryOptions?: MessagingExtensionQueryOptions;
  /**
   * @member {string} [state] State parameter passed back to the bot after
   * authentication/configuration flow
   */
  state?: string;
}

/**
 * @interface
 * An interface representing MessagingExtensionAttachment.
 * Messaging extension attachment.
 *
 * @extends builder.Attachment
 */
export interface MessagingExtensionAttachment extends builder.Attachment {
  /**
   * @member {Attachment} [preview]
   */
  preview?: builder.Attachment;
}

/**
 * @interface
 * An interface representing MessagingExtensionSuggestedAction.
 * Messaging extension Actions (Only when type is auth or config)
 *
 */
export interface MessagingExtensionSuggestedAction {
  /**
   * @member {CardAction[]} [actions] Actions
   */
  actions?: builder.CardAction[];
}

/**
 * @interface
 * An interface representing MessagingExtensionResult.
 * Messaging extension result
 *
 */
export interface MessagingExtensionResult {
  /**
   * @member {AttachmentLayout} [attachmentLayout] Hint for how to deal with
   * multiple attachments. Possible values include: 'list', 'grid'
   */
  attachmentLayout?: AttachmentLayout;
  /**
   * @member {Type2} [type] The type of the result. Possible values include:
   * 'result', 'auth', 'config', 'message'
   */
  type?: Type2;
  /**
   * @member {MessagingExtensionAttachment[]} [attachments] (Only when type is
   * result) Attachments
   */
  attachments?: MessagingExtensionAttachment[];
  /**
   * @member {MessagingExtensionSuggestedAction} [suggestedActions]
   */
  suggestedActions?: MessagingExtensionSuggestedAction;
  /**
   * @member {string} [text] (Only when type is message) Text
   */
  text?: string;
}

/**
 * @interface
 * An interface representing MessagingExtensionResponse.
 * Messaging extension response
 *
 */
export interface MessagingExtensionResponse {
  /**
   * @member {MessagingExtensionResult} [composeExtension]
   */
  composeExtension?: MessagingExtensionResult;
}

/**
 * @interface
 * An interface representing FileConsentCard.
 * File consent card attachment.
 *
 */
export interface FileConsentCard {
  /**
   * @member {string} [description] File description.
   */
  description?: string;
  /**
   * @member {number} [sizeInBytes] Size of the file to be uploaded in Bytes.
   */
  sizeInBytes?: number;
  /**
   * @member {any} [acceptContext] Context sent back to the Bot if user
   * consented to upload. This is free flow schema and is sent back in Value
   * field of Activity.
   */
  acceptContext?: any;
  /**
   * @member {any} [declineContext] Context sent back to the Bot if user
   * declined. This is free flow schema and is sent back in Value field of
   * Activity.
   */
  declineContext?: any;
}

/**
 * @interface
 * An interface representing FileDownloadInfo.
 * File download info attachment.
 *
 */
export interface FileDownloadInfo {
  /**
   * @member {string} [downloadUrl] File download url.
   */
  downloadUrl?: string;
  /**
   * @member {string} [uniqueId] Unique Id for the file.
   */
  uniqueId?: string;
  /**
   * @member {string} [fileType] Type of file.
   */
  fileType?: string;
  /**
   * @member {any} [etag] ETag for the file.
   */
  etag?: any;
}

/**
 * @interface
 * An interface representing FileInfoCard.
 * File info card.
 *
 */
export interface FileInfoCard {
  /**
   * @member {string} [uniqueId] Unique Id for the file.
   */
  uniqueId?: string;
  /**
   * @member {string} [fileType] Type of file.
   */
  fileType?: string;
  /**
   * @member {any} [etag] ETag for the file.
   */
  etag?: any;
}

/**
 * @interface
 * An interface representing FileUploadInfo.
 * Information about the file to be uploaded.
 *
 */
export interface FileUploadInfo {
  /**
   * @member {string} [name] Name of the file.
   */
  name?: string;
  /**
   * @member {string} [uploadUrl] URL to an upload session that the bot can use
   * to set the file contents.
   */
  uploadUrl?: string;
  /**
   * @member {string} [contentUrl] URL to file.
   */
  contentUrl?: string;
  /**
   * @member {string} [uniqueId] ID that uniquely identifies the file.
   */
  uniqueId?: string;
  /**
   * @member {string} [fileType] Type of the file.
   */
  fileType?: string;
}

/**
 * @interface
 * An interface representing FileConsentCardResponse.
 * Represents the value of the invoke activity sent when the user acts on a
 * file consent card
 *
 */
export interface FileConsentCardResponse {
  /**
   * @member {Action} [action] The action the user took. Possible values
   * include: 'accept', 'decline'
   */
  action?: Action;
  /**
   * @member {any} [context] The context associated with the action.
   */
  context?: any;
  /**
   * @member {FileUploadInfo} [uploadInfo] If the user accepted the file,
   * contains information about the file to be uploaded.
   */
  uploadInfo?: FileUploadInfo;
}

/**
 * @interface
 * An interface representing TeamsConnectorClientOptions.
 * @extends ServiceClientOptions
 */
export interface TeamsConnectorClientOptions extends ServiceClientOptions {
  /**
   * @member {string} [baseUri]
   */
  baseUri?: string;
}

/**
 * Defines values for Type.
 * Possible values include: 'ViewAction', 'OpenUri', 'HttpPOST', 'ActionCard'
 * @readonly
 * @enum {string}
 */
export type Type = 'ViewAction' | 'OpenUri' | 'HttpPOST' | 'ActionCard';

/**
 * Defines values for ActivityImageType.
 * Possible values include: 'avatar', 'article'
 * @readonly
 * @enum {string}
 */
export type ActivityImageType = 'avatar' | 'article';

/**
 * Defines values for Os.
 * Possible values include: 'default', 'iOS', 'android', 'windows'
 * @readonly
 * @enum {string}
 */
export type Os = 'default' | 'iOS' | 'android' | 'windows';

/**
 * Defines values for Type1.
 * Possible values include: 'textInput', 'dateInput', 'multichoiceInput'
 * @readonly
 * @enum {string}
 */
export type Type1 = 'textInput' | 'dateInput' | 'multichoiceInput';

/**
 * Defines values for Style.
 * Possible values include: 'compact', 'expanded'
 * @readonly
 * @enum {string}
 */
export type Style = 'compact' | 'expanded';

/**
 * Defines values for AttachmentLayout.
 * Possible values include: 'list', 'grid'
 * @readonly
 * @enum {string}
 */
export type AttachmentLayout = 'list' | 'grid';

/**
 * Defines values for Type2.
 * Possible values include: 'result', 'auth', 'config', 'message'
 * @readonly
 * @enum {string}
 */
export type Type2 = 'result' | 'auth' | 'config' | 'message';

/**
 * Defines values for Action.
 * Possible values include: 'accept', 'decline'
 * @readonly
 * @enum {string}
 */
export type Action = 'accept' | 'decline';

/**
 * Contains response data for the fetchChannelList operation.
 */
export type TeamsFetchChannelListResponse = ConversationList & {
  /**
   * The underlying HTTP response.
   */
  _response: msRest.HttpResponse & {
      /**
       * The response body as text (string format)
       */
      bodyAsText: string;
      /**
       * The response body as parsed JSON or XML
       */
      parsedBody: ConversationList;
    };
};

/**
 * Contains response data for the fetchTeamDetails operation.
 */
export type TeamsFetchTeamDetailsResponse = TeamDetails & {
  /**
   * The underlying HTTP response.
   */
  _response: msRest.HttpResponse & {
      /**
       * The response body as text (string format)
       */
      bodyAsText: string;
      /**
       * The response body as parsed JSON or XML
       */
      parsedBody: TeamDetails;
    };
};