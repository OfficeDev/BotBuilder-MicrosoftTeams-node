import * as builder from 'botbuilder';
import { TeamsConnectorClient } from './';
import * as models from './schema/models';

export class TeamsContext {
  public static readonly stateKey: symbol = Symbol('TeamsContextCacheKey');

  constructor (
    public readonly turnContext: builder.TurnContext,
    public readonly teamsConnectorClient: TeamsConnectorClient
  ) {
  }

  public static from(turnContext: builder.TurnContext): TeamsContext {
    return turnContext.turnState.get(TeamsContext.stateKey);
  }

  public get eventType(): string {
    return this.getTeamsChannelData().eventType;
  }

  public get team(): models.TeamInfo {
    return this.getTeamsChannelData().team;
  }

  public get channel(): models.ChannelInfo {
    return this.getTeamsChannelData().channel;
  }

  public get tenant(): models.TenantInfo {
    return this.getTeamsChannelData().tenant;
  }

  public getGeneralChannel(): models.ChannelInfo {
    const channelData = this.getTeamsChannelData();
    if (channelData && channelData.team) {
      return {
        id: channelData.team.id
      };
    }
    throw new Error('Failed to process channel data in Activity. ChannelData is missing Team property.');
  }

  public getTeamsChannelData(): models.TeamsChannelData {
    const channelData = this.turnContext.activity.channelData;
    if (!channelData) {
      throw new Error('ChannelData missing Activity');
    } else {
      return channelData as models.TeamsChannelData;
    }
  }

  public getActivityTextWithoutMentions(): string {
    const activity = this.turnContext.activity;
    if (activity.entities && activity.entities.length === 0) {
      return activity.text;
    }
    
    const recvBotId = activity.recipient.id;
    const recvBotMentioned = <builder.Mention[]> activity.entities.filter(e =>
      (e.type === 'mention') && (e as builder.Mention).mentioned.id === recvBotId
    );

    if (recvBotMentioned.length === 0) {
      return activity.text;
    }

    let strippedText = activity.text;
    recvBotMentioned.forEach(m => m.text && (strippedText = strippedText.replace(m.text, '')));
    return strippedText.trim();
  }

  public getConversationParametersForCreateOrGetDirectConversation(user: builder.ChannelAccount): builder.ConversationParameters {
    const channelData = {} as  models.TeamsChannelData;
    if (this.tenant && this.tenant.id) {
      channelData.tenant = {
        id: this.tenant.id
      };
    }

    return <builder.ConversationParameters> {
      bot: this.turnContext.activity.recipient,
      channelData,
      members: [user]
    };
  }

  // [TODO] need to resolve schema problems in botframework-connector where Activity only supports Entity but not Mention.
  // (entity will be dropped out before sending out due to schema checking)
  public static addMentionToText<T extends builder.Activity>(activity: T, mentionedEntity: builder.ChannelAccount, mentionText?: string): T {
    if (!mentionedEntity || !mentionedEntity.id) {
      throw new Error('Mentioned entity and entity ID cannot be null');
    }

    if (!mentionedEntity.name && !mentionText) {
      throw new Error('Either mentioned name or mentionText must have a value');
    }

    if (!!mentionText) {
      mentionedEntity.name = mentionText;
    }

    const mentionEntityText = `<at>${mentionedEntity.name}</at>`;
    activity.text += ` ${mentionEntityText}`;
    activity.entities = [
      ...(activity.entities || []),
      {
        type: 'mention',
        text: mentionEntityText,
        mentioned: mentionedEntity
      } as builder.Mention
    ];

    return activity;
  }

  public static notifyUser<T extends builder.Activity>(activity: T): T {
    let channelData: models.TeamsChannelData = activity.channelData || {};
    channelData.notification = {
      alert: true
    };
    activity.channelData = channelData;
    return activity;
  }

  public static isTeamsChannelAccount(channelAccount: builder.ChannelAccount | any[]): channelAccount is models.TeamsChannelAccount {
    const o = channelAccount as models.TeamsChannelAccount;
    return !!o
        &&(!!o.id && !!o.name)
        && (!!o.aadObjectId || !!o.email || !!o.givenName || !!o.surname || !!o.userPrincipalName);
  }

  public static isTeamsChannelAccounts(channelAccount: builder.ChannelAccount[] | any[]): channelAccount is models.TeamsChannelAccount[] {
    return Array.isArray(channelAccount) && channelAccount.every(x => this.isTeamsChannelAccount(x));
  }
}
