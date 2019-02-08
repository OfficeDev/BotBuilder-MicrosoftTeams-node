import * as builder from 'botbuilder';

export class DropChatActivitiesMiddleware implements builder.Middleware {
  public async onTurn(context: builder.TurnContext, next: () => Promise<void>): Promise<void> {
    const convType = context.activity.conversation && context.activity.conversation.conversationType;
    const isChat = !!convType && (convType === 'personal' || convType == 'groupChat');
    // only non-channel activities can pass through
    if (!isChat) {
      return next();
    }
  }
}
