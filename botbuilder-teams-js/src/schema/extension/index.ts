import * as builder from 'botbuilder';
import * as teams from '../models';

/**
 * Defines values for Type.
 * Possible values include: 'ViewAction', 'OpenUri', 'HttpPOST', 'ActionCard'
 * @readonly
 * @enum {string}
 */
export type O365ConnectorCardActionType = 'ViewAction' | 'OpenUri' | 'HttpPOST' | 'ActionCard';

/**
 * @interface
 * An interface representing O365ConnectorCardActionBase.
 * O365 connector card action base
 *
 */
export interface O365ConnectorCardActionBase {
  /**
   * @member {Type} [type] Type of the action. Possible values include:
   * 'ViewAction', 'OpenUri', 'HttpPOST', 'ActionCard'
   */
  '@type'?: O365ConnectorCardActionType;
  /**
   * @member {string} [name] Name of the action that will be used as button
   * title
   */
  name?: string;
  /**
   * @member {string} [id] Action Id
   */
  '@id'?: string;
}

/**
 * Defines values for Type1.
 * Possible values include: 'textInput', 'dateInput', 'multichoiceInput'
 * @readonly
 * @enum {string}
 */
export type O365ConnectorCardInputType = 'textInput' | 'dateInput' | 'multichoiceInput';

/**
 * @interface
 * An interface representing O365ConnectorCardInputBase.
 * O365 connector card input for ActionCard action
 *
 */
export interface O365ConnectorCardInputBase {
  /**
   * @member {Type1} [type] Input type name. Possible values include:
   * 'textInput', 'dateInput', 'multichoiceInput'
   */
  '@type'?: O365ConnectorCardInputType;
  /**
   * @member {string} [id] Input Id. It must be unique per entire O365
   * connector card.
   */
  id?: string;
  /**
   * @member {boolean} [isRequired] Define if this input is a required field.
   * Default value is false.
   */
  isRequired?: boolean;
  /**
   * @member {string} [title] Input title that will be shown as the placeholder
   */
  title?: string;
  /**
   * @member {string} [value] Default value for this input field
   */
  value?: string;
}

export interface TeamsAttachment<ContentType> extends builder.Attachment {
  content: ContentType;
}

export type FileDownloadInfoAttachment = TeamsAttachment<teams.FileDownloadInfo>;
