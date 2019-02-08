import * as builder from 'botbuilder';
import { TeamsChannelData } from '../schema/models';

export class TeamSpecificConversationState extends builder.BotState {
  constructor (storage: builder.Storage) {
    const storageKeyGenerator = (turnContext: builder.TurnContext): Promise<string> => {
      const teamsChannelData: TeamsChannelData = turnContext.activity.channelData;
      if (teamsChannelData.team && teamsChannelData.team.id) {
        return Promise.resolve(`team/${turnContext.activity.channelId}/${teamsChannelData.team.id}`);
      } else {
        return Promise.resolve(`chat/${turnContext.activity.channelId}/${turnContext.activity.conversation.id}`);
      }
    };
    super(storage, storageKeyGenerator);
  }
}