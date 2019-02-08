import * as builder from 'botbuilder';
import { MicrosoftAppCredentials } from 'botframework-connector';
import { TeamsConnectorClient, TeamsContext } from '../';

export class TeamsMiddleware implements builder.Middleware {
  public async onTurn(context: builder.TurnContext, next: () => Promise<void>): Promise<void> {
    const credentials: MicrosoftAppCredentials = context.adapter && (context.adapter as any).credentials;
    if (context.activity.channelId === 'msteams' && !!credentials) {
      const serviceUrl = context.activity && context.activity.serviceUrl;
      const teamsConnectorClient = new TeamsConnectorClient(credentials, { baseUri: serviceUrl });
      context.turnState.set(TeamsContext.stateKey, new TeamsContext(context, teamsConnectorClient));
    }
    return next();
  }
}
