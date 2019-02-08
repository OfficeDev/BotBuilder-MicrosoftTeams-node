import { BotFrameworkAdapter, ConversationReference, TurnContext, ConversationParameters, Activity, ConversationAccount, ResourceResponse, ActivityTypes, ConversationsGetConversationMembersResponse, ConversationsGetActivityMembersResponse } from 'botbuilder';
import { ConnectorClient } from 'botframework-connector';
import { TeamsChannelData, TeamsChannelAccount } from './schema';
import { TeamsContext } from './teamsContext';

export class TeamsAdapter extends BotFrameworkAdapter {
  /**
   * Lists the members of a given activity as specified in a TurnContext.
   *
   * @remarks
   * Returns an array of ChannelAccount objects representing the users involved in a given activity.
   *
   * This is different from `getConversationMembers()` in that it will return only those users
   * directly involved in the activity, not all members of the conversation.
   * @param context Context for the current turn of conversation with the user.
   * @param activityId (Optional) activity ID to enumerate. If not specified the current activities ID will be used.
   */
  public async getActivityMembers(context: TurnContext, activityId?: string): Promise<TeamsChannelAccount[]> {
    const members = await super.getActivityMembers(context)
      .then((res: ConversationsGetActivityMembersResponse) => 
        this.merge_objectId_and_aadObjectId(JSON.parse(res._response.bodyAsText)));
    if (TeamsContext.isTeamsChannelAccounts(members)) {
      return members;
    } else {
      throw new Error('Members are not TeamsChannelAccount[]');
    }
  }
  
  /**
   * Lists the members of the current conversation as specified in a TurnContext.
   *
   * @remarks
   * Returns an array of ChannelAccount objects representing the users currently involved in the conversation
   * in which an activity occured.
   *
   * This is different from `getActivityMembers()` in that it will return all
   * members of the conversation, not just those directly involved in the activity.
   * @param context Context for the current turn of conversation with the user.
   */
  public async getConversationMembers(context: TurnContext): Promise<TeamsChannelAccount[]> {
    const members = await super.getConversationMembers(context)
      .then((res: ConversationsGetConversationMembersResponse) =>
        this.merge_objectId_and_aadObjectId(JSON.parse(res._response.bodyAsText)));      
    if (TeamsContext.isTeamsChannelAccounts(members)) {
      return members;
    } else {
      throw new Error('Members are not TeamsChannelAccount[]');
    }
  }

  /**
   * Starts a new conversation with a user. This is typically used to Direct Message (DM) a member
   * of a group.
   *
   * @remarks
   * This function creates a new conversation between the bot and a single user, as specified by
   * the ConversationReference passed in. In multi-user chat environments, this typically means
   * starting a 1:1 direct message conversation with a single user. If called on a reference
   * already representing a 1:1 conversation, the new conversation will continue to be 1:1.
   *
   * * In order to use this method, a ConversationReference must first be extracted from an incoming
   * activity. This reference can be stored in a database and used to resume the conversation at a later time.
   * The reference can be created from any incoming activity using `TurnContext.getConversationReference(context.activity)`.
   *
   * The processing steps for this method are very similar to [processActivity()](#processactivity)
   * in that a `TurnContext` will be created which is then routed through the adapters middleware
   * before calling the passed in logic handler. The key difference is that since an activity
   * wasn't actually received from outside, it has to be created by the bot.  The created activity will have its address
   * related fields populated but will have a `context.activity.type === undefined`..
   *
   * ```JavaScript
   * // Get group members conversation reference
   * const reference = TurnContext.getConversationReference(context.activity);
   *
   * // Start a new conversation with the user
   * await adapter.createConversation(reference, async (ctx) => {
   *    await ctx.sendActivity(`Hi (in private)`);
   * });
   * ```
   * @param reference A `ConversationReference` of the user to start a new conversation with.
   * @param logic A function handler that will be called to perform the bot's logic after the the adapters middleware has been run.
   */
  public async createTeamsConversation(reference: Partial<ConversationReference>, tenantIdOrTurnContext: string | TurnContext, logic?: (context: TurnContext) => Promise<void>): Promise<void> {
    if (!reference.serviceUrl) { throw new Error(`TeamsAdapter.createTeamsConversation(): missing serviceUrl.`); }

    // Create conversation
    const tenantId: string = (tenantIdOrTurnContext instanceof TurnContext)
      ? TeamsContext.from(tenantIdOrTurnContext).tenant.id || undefined
      : <string> tenantIdOrTurnContext;
    const channelData: TeamsChannelData = { tenant: {id: tenantId} };
    const parameters: ConversationParameters = { bot: reference.bot, members: [reference.user], channelData } as ConversationParameters;
    const client: ConnectorClient = this.createConnectorClient(reference.serviceUrl);
    const response = await client.conversations.createConversation(parameters);

    // Initialize request and copy over new conversation ID and updated serviceUrl.
    const request: Partial<Activity> = TurnContext.applyConversationReference(
      { type: 'event', name: 'createConversation' },
      reference,
      true
    );
    request.conversation = { id: response.id } as ConversationAccount;
    if (response.serviceUrl) { request.serviceUrl = response.serviceUrl; }

    // Create context and run middleware
    const context: TurnContext = this.createContext(request);
    await this.runMiddleware(context, logic as any);
  }

  /**
   * Sends a set of activities to the user. An array of responses from the server will be returned.
   *
   * @remarks
   * Prior to delivery, the activities will be updated with information from the `ConversationReference`
   * for the contexts [activity](#activity) and if any activities `type` field hasn't been set it will be
   * set to a type of `message`. The array of activities will then be routed through any [onSendActivities()](#onsendactivities)
   * handlers before being passed to `adapter.sendActivities()`.
   *
   * ```JavaScript
   * await context.sendActivities([
    *    { type: 'typing' },
    *    { type: 'delay', value: 2000 },
    *    { type: 'message', text: 'Hello... How are you?' }
    * ]);
    * ```
    * @param activities One or more activities to send to the user.
    */
    public createReplyChain(turnContext: TurnContext, activities: Partial<Activity>[], inGeneralChannel?: boolean): Promise<ResourceResponse[]> {
      let sentNonTraceActivity: boolean = false;
      const teamsCtx = TeamsContext.from(turnContext);
      const ref: Partial<ConversationReference> = TurnContext.getConversationReference(turnContext.activity);
      const output: Partial<Activity>[] = activities.map((a: Partial<Activity>) => {
        const o: Partial<Activity> = TurnContext.applyConversationReference({...a}, ref);
        try {
          o.conversation.id = inGeneralChannel
            ? teamsCtx.getGeneralChannel().id
            : teamsCtx.channel.id;
        } catch (e) {
          // do nothing for fields fetching error
        }
        if (!o.type) { o.type = ActivityTypes.Message; }
        if (o.type !== ActivityTypes.Trace) { sentNonTraceActivity = true; }
        return o;
      });

      return turnContext['emit'](turnContext['_onSendActivities'], output, () => {
        return super.sendActivities(turnContext, output)
          .then((responses: ResourceResponse[]) => {
            // Set responded flag
            if (sentNonTraceActivity) { turnContext.responded = true; }
            return responses;
          });
      });
    }

    /**
     * SMBA sometimes returns "objectId" and sometimes returns "aadObjectId".
     * Use this function to unify them into "aadObjectId" that is defined by schema.
     * @param members raw members array
     */
    private merge_objectId_and_aadObjectId(members: any[]) {
      if (members) {
        members.forEach(m => {
          if (!m.aadObjectId && m.objectId) {
            m.aadObjectId = m.objectId;
          }
          delete m.objectId;
        });
      }
      return members;
    }
}