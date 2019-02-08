import * as models from '../schema/models';
import { TurnContext, ChannelAccount } from 'botbuilder';

export type TeamEventType =
    'teamMembersAdded' 
  | 'teamMembersRemoved'
  | 'channelCreated'
  | 'channelDeleted'
  | 'channelRenamed'
  | 'teamRenamed';

export interface ITeamEvent
{
    readonly eventType: TeamEventType;
    readonly team: models.TeamInfo;
    readonly tenant: models.TenantInfo;
    readonly turnContext: TurnContext;
}

export interface IChannelCreatedEvent extends ITeamEvent {
  readonly eventType: 'channelCreated';
  readonly channel: models.ChannelInfo
}

export interface IChannelDeletedEvent extends ITeamEvent {
  readonly eventType: 'channelDeleted';
  readonly channel: models.ChannelInfo;
}

export interface IChannelRenamedEvent extends ITeamEvent {
  readonly eventType: 'channelRenamed';
  readonly channel: models.ChannelInfo;
}

export interface IMembersAddedEvent extends ITeamEvent {
  readonly eventType: 'teamMembersAdded';
  readonly membersAdded: ChannelAccount[];
}

export interface IMembersRemovedEvent extends ITeamEvent {
  readonly eventType: 'teamMembersRemoved';
  readonly membersRemoved: ChannelAccount[];
}

export interface ITeamRenamedEvent extends ITeamEvent {
  readonly eventType: 'teamRenamed';
}
