import * as builder from 'botbuilder';

export class DropNonTeamsActivitiesMiddleware implements builder.Middleware {
  public async onTurn(context: builder.TurnContext, next: () => Promise<void>): Promise<void> {
    // only 'msteams' activities can pass through
    if (context.activity.channelId === 'msteams') {
      return next();
    }
  }
}
