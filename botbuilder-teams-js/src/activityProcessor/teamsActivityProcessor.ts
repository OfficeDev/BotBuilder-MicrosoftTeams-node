import { TurnContext, ActivityTypes } from 'botbuilder';
import { TeamsChannelData, FileDownloadInfo, FileDownloadInfoAttachment } from '../schema';
import { TeamEventType, IMembersAddedEvent, IMembersRemovedEvent, IChannelCreatedEvent, IChannelDeletedEvent, IChannelRenamedEvent, ITeamRenamedEvent } from './teamsEvents';
import { ITeamsInvokeActivityHandler, InvokeActivity } from './teamsInvoke';
import { TeamsFactory } from './teamsFactory';

export interface IMessageActivityHandler {
  onMessage?: (context: TurnContext) => Promise<void>;
  onMessageWithFileDownloadInfo?: (context: TurnContext, attachment: FileDownloadInfo) => Promise<void>;
}

export interface ITeamsConversationUpdateActivityHandler {
  onTeamMembersAdded?: (event: IMembersAddedEvent) => Promise<void>;
  onTeamMembersRemoved?: (event: IMembersRemovedEvent) => Promise<void>;
  onChannelCreated?: (event: IChannelCreatedEvent) => Promise<void>;
  onChannelDeleted?: (event: IChannelDeletedEvent) => Promise<void>;
  onChannelRenamed?: (event: IChannelRenamedEvent) => Promise<void>;
  onTeamRenamed?: (event: ITeamRenamedEvent) => Promise<void>;
  onConversationUpdateActivity?: (turnContext: TurnContext) => Promise<void>;
}

export interface IMessageReactionActivityHandler {
  onMessageReaction?: (turnContext: TurnContext) => Promise<void>;
}

export enum ActivityTypesEx {
  InvokeResponse = 'invokeResponse'
}

export class TeamsActivityProcessor {
  constructor (
    public messageActivityHandler?: IMessageActivityHandler,
    public conversationUpdateActivityHandler?: ITeamsConversationUpdateActivityHandler,
    public invokeActivityHandler?: ITeamsInvokeActivityHandler,
    public messageReactionActivityHandler?: IMessageReactionActivityHandler
  ) {
  }

  public async processIncomingActivity (turnContext: TurnContext) {
    switch (turnContext.activity.type) {
      case ActivityTypes.Message:
        {
          await this.processOnMessageActivity(turnContext);
          break;
        }

      case ActivityTypes.ConversationUpdate:
        {
          await this.processTeamsConversationUpdate(turnContext);
          break;            
        }
      
      case ActivityTypes.Invoke:
        {
          const invokeResponse = await InvokeActivity.dispatchHandler(this.invokeActivityHandler, turnContext);
          if (invokeResponse) {
            await turnContext.sendActivity({ value: invokeResponse, type: ActivityTypesEx.InvokeResponse });
          }
          break;
        }

      case ActivityTypes.MessageReaction:
        {
          const handler = this.messageReactionActivityHandler;
          if (handler && handler.onMessageReaction) {
            await handler.onMessageReaction(turnContext);
          }
          break;
        }
    }
  }

  private async processOnMessageActivity (turnContext: TurnContext) {
    const handler = this.messageActivityHandler;
    if (handler) {
      if (handler.onMessageWithFileDownloadInfo) {
        const attachments = turnContext.activity.attachments || [];
        const fileDownload = attachments.map(x => TeamsFactory.isFileDownloadInfoAttachment(x) && x).shift();
        if (fileDownload) {
          return await handler.onMessageWithFileDownloadInfo(turnContext, fileDownload.content);
        }
      }

      if (handler.onMessage) {
        await handler.onMessage(turnContext);
      }
    }
  }

  private async processTeamsConversationUpdate (turnContext: TurnContext) {
    const handler = this.conversationUpdateActivityHandler;
    if (handler) {
      const channelData: TeamsChannelData = turnContext.activity.channelData;
      if (channelData && channelData.eventType) {
        switch (channelData.eventType as TeamEventType) {
          case 'teamMembersAdded':
            handler.onTeamMembersAdded && await handler.onTeamMembersAdded({
              eventType: 'teamMembersAdded',
              turnContext: turnContext,
              team: channelData.team,
              tenant: channelData.tenant,
              membersAdded: turnContext.activity.membersAdded,
            });
            break;

          case 'teamMembersRemoved':
            handler.onTeamMembersRemoved && await handler.onTeamMembersRemoved({
              eventType: 'teamMembersRemoved',
              turnContext: turnContext,
              team: channelData.team,
              tenant: channelData.tenant,
              membersRemoved: turnContext.activity.membersRemoved
            });
            break;

          case 'channelCreated':
            handler.onChannelCreated && await handler.onChannelCreated({
              eventType: 'channelCreated',
              turnContext: turnContext,
              team: channelData.team,
              tenant: channelData.tenant,
              channel: channelData.channel
            });
            break;

          case 'channelDeleted':
            handler.onChannelDeleted && await handler.onChannelDeleted({
              eventType: 'channelDeleted',
              turnContext: turnContext,
              team: channelData.team,
              tenant: channelData.tenant,
              channel: channelData.channel
            });
            break;

          case 'channelRenamed':
            handler.onChannelRenamed && await handler.onChannelRenamed({
              eventType: 'channelRenamed',
              turnContext: turnContext,
              team: channelData.team,
              tenant: channelData.tenant,
              channel: channelData.channel
            });
            break;

          case 'teamRenamed':
            handler.onTeamRenamed && await handler.onTeamRenamed({
              eventType: 'teamRenamed',
              turnContext: turnContext,
              team: channelData.team,
              tenant: channelData.tenant
            });
            break;
        }  
      }

      if (handler.onConversationUpdateActivity) {
        await handler.onConversationUpdateActivity(turnContext);
      }
    }
  }
}