import * as builder from 'botbuilder';
import { TeamsChannelData } from '../schema/models';

export class DropChannelActivitiesMiddleware implements builder.Middleware {
  public async onTurn(context: builder.TurnContext, next: () => Promise<void>): Promise<void> {
    // only non-channel activities can pass through
    const teamsChannelData: TeamsChannelData = context.activity.channelData;
    if (!teamsChannelData.team) {
      return next();
    }
  }
}
