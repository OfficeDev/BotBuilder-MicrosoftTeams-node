// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { StatePropertyAccessor, TurnContext, CardFactory, BotState, Activity } from 'botbuilder';
import * as teams from 'botbuilder-teams-js';
import * as request from 'request';
import * as fs from 'fs';

// Turn counter property
const TURN_COUNTER = 'turnCounterProperty';

export class EchoBot {
    
    private readonly countAccessor: StatePropertyAccessor<number>;
    private readonly conversationState: BotState;

    /**
     * 
     * @param {ConversationState} conversation state object
     */
    constructor (conversationState: BotState) {
        // Create a new state accessor property. See https://aka.ms/about-bot-state-accessors to learn more about the bot state and state accessors.        
        this.countAccessor = conversationState.createProperty(TURN_COUNTER);
        this.conversationState = conversationState;
    }

    /**
     * Use onTurn to handle an incoming activity, received from a user, process it, and reply as needed
     * 
     * @param {TurnContext} context on turn context object.
     */
    async onTurn(turnContext: TurnContext) {
        const activityProc = new teams.TeamsActivityProcessor();

        activityProc.messageActivityHandler = {
            onMessageWithFileDownloadInfo: async (ctx, file) => {
                await turnContext.sendActivity({ textFormat: 'xml', text: `<b>Received File</b> <pre>${JSON.stringify(file, null, 2)}</pre>`});
            },

            onMessage: async (ctx) => {
                const teamsCtx = teams.TeamsContext.from(turnContext);
                const text = teamsCtx.getActivityTextWithoutMentions();
                const adapter = turnContext.adapter as teams.TeamsAdapter;

                switch (text.toLowerCase()) {
                    case 'cards':
                        await this.sentCards(ctx);
                        break;

                    case 'reply-chain':
                        await adapter.createReplyChain(ctx, [{ text: 'New reply chain' }]);
                        break;

                    case '1:1':
                        // create 1:1 conversation
                        const tenantId = teamsCtx.tenant.id;
                        const ref = TurnContext.getConversationReference(ctx.activity);        
                        await adapter.createTeamsConversation(ref, tenantId, async (newCtx) => {
                            await newCtx.sendActivity(`Hi (in private)`);
                        });
                        break;
                    
                    case 'members':
                        const members = await adapter.getConversationMembers(turnContext);
                        const actMembers = await adapter.getActivityMembers(turnContext);
                        await ctx.sendActivity({ 
                            textFormat: 'xml', 
                            text: `
                                <b>Activity Members</b></br>
                                <pre>${JSON.stringify(actMembers, null, 2)}</pre>
                                <b>Conversation Members</b></br>
                                <pre>${JSON.stringify(members, null, 2)}</pre>
                        `});
                        break;
                    
                    case 'team-info':
                        const teamId = teamsCtx.team.id;
                        const chList = await teamsCtx.teamsConnectorClient.teams.fetchChannelList(teamId);
                        const tmDetails = await teamsCtx.teamsConnectorClient.teams.fetchTeamDetails(teamId);
                        await turnContext.sendActivity({ textFormat: 'xml', text: `<pre>${JSON.stringify(chList, null, 2)}</pre>`});
                        await turnContext.sendActivity({ textFormat: 'xml', text: `<pre>${JSON.stringify(tmDetails, null, 2)}</pre>`});
                        break;
                    
                    default:
                        let count = await this.countAccessor.get(turnContext);
                        count = count === undefined ? 1 : ++count;
                        await this.countAccessor.set(turnContext, count);

                        let activity: Partial<Activity> = {
                            textFormat: 'xml',
                            text: `${ count }: You said "${ turnContext.activity.text }"`
                        };

                        // activity = teamsCtx.addMentionToText(activity as Activity, turnContext.activity.from);
                        activity = teams.TeamsContext.notifyUser(activity as Activity);

                        await ctx.sendActivity(activity);
                        await this.conversationState.saveChanges(turnContext);
                }            
            }
        };
       
        activityProc.conversationUpdateActivityHandler = {
            onChannelRenamed: async (event) => {
                const e = {
                    channel: event.channel,
                    eventType: event.eventType,
                    team: event.team,
                    tenant: event.tenant
                }
                event.turnContext.sendActivity({ textFormat: 'xml', text: `[Conversation Update] <pre>${JSON.stringify(e, null, 2)}</pre>`})
            }
        };

        activityProc.messageReactionActivityHandler = {
            onMessageReaction: async (turnContext) => {
                const added = turnContext.activity.reactionsAdded;
                const removed = turnContext.activity.reactionsRemoved;
                await turnContext.sendActivity({ 
                    textFormat: 'xml', 
                    text: `
                        <h1><b>[Reaction Event]</b></h1></br>
                        <b>Added</b></br>
                        <pre>${JSON.stringify(added, null, 2)}</pre>
                        <b>Removed</b></br>
                        <pre>${JSON.stringify(removed, null, 2)}</pre>
                        <b>Activity</b></br>
                        <pre>${JSON.stringify(turnContext.activity, null, 2)}</pre>
                `});

                const adapter = turnContext.adapter as teams.TeamsAdapter;
                const members = await adapter.getConversationMembers(turnContext);
                const fromAad = turnContext.activity.from.aadObjectId;
                const member = members.find(m => m.aadObjectId === fromAad);
                const memberName = member && member.givenName;
                const isLike = added && added[0] && added[0].type === 'like';
                if (memberName) {
                    const text = isLike
                        ? `<b>${memberName}</b>, thanks for liking my message! 😍😘`
                        : `<b>${memberName}</b>, why don't you like what I said? 😭😢`;
                    await turnContext.sendActivity({ textFormat: 'xml', text });
                }
            }
        };

        activityProc.invokeActivityHandler = {
            onO365CardAction: async (ctx: TurnContext, query: teams.O365ConnectorCardActionQuery) => {
                let userName = ctx.activity.from.name;
                let body = JSON.parse(query.body);
                let msg: Partial<Activity> = {
                    summary: 'Thanks for your input!',
                    textFormat: 'xml',
                    text: `<h2>Thanks, ${userName}!</h2><br/><h3>Your input action ID:</h3><br/><pre>${query.actionId}</pre><br/><h3>Your input body:</h3><br/><pre>${JSON.stringify(body, null, 2)}</pre>`
                };
                await turnContext.sendActivity(msg);
                return { status: 200 };
            },
            
            onMessagingExtensionQuery: async (ctx: TurnContext, query: teams.MessagingExtensionQuery) => {
                type R = teams.InvokeResponseTypeOf<'onMessagingExtensionQuery'>;

                let preview = CardFactory.thumbnailCard('Search Item Card', 'This is to show the search result');
                let heroCard = CardFactory.heroCard('Result Card', '<pre>This card mocks the CE results</pre>');

                return Promise.resolve(<R> {
                    status: 200,
                    'body': {
                        'composeExtension': {
                            'type': 'result',
                            'attachmentLayout': 'list',
                            'attachments': [
                                { ...heroCard, preview }
                            ]
                        }
                    }
                });
            },

            onFileConsent: async (ctx: TurnContext, query: teams.FileConsentCardResponse) => {
                await turnContext.sendActivity({ textFormat: 'xml', text: `[onFileConsent] <pre>${JSON.stringify(query, null, 2)}</pre>`});
                if (query.action === 'accept') {
                    const fname = 'teamsAppManifest/icon-color.png';
                    const fileInfo = fs.statSync(fname);
                    const file = new Buffer(fs.readFileSync(fname, 'binary'), 'binary');
                    const r = await new Promise((resolve, reject) => {
                        request.put({
                            uri: query.uploadInfo.uploadUrl,
                            headers: {
                               'Content-Length': fileInfo.size,
                               'Content-Range': `bytes 0-${fileInfo.size-1}/${fileInfo.size}`
                            },
                            encoding: null,
                            body: file, 
                            json: true
                        }, async (err, res) => {
                            const downloadCard = teams.TeamsFactory.fileInfoCard(
                                query.uploadInfo.name,
                                query.uploadInfo.contentUrl,
                                {
                                    'uniqueId': query.uploadInfo.uniqueId,
                                    'fileType': query.uploadInfo.fileType
                                }
                            );
                            await ctx.sendActivity({ attachments: [downloadCard] });
                            err ? resolve(err) : resolve(res.body);
                        });
                    });
                    await turnContext.sendActivity({ textFormat: 'xml', text: `[file upload] <pre>${JSON.stringify(r, null, 2)}</pre>`});
                }
                return { status: 200 };
            },

            onInvoke: async (ctx) => {
                await turnContext.sendActivity({ textFormat: 'xml', text: `[General onInvoke] <pre>${JSON.stringify(ctx.activity, null, 2)}</pre>`});
                return { status: 200, body: { composeExtensions: {} } };
            }
        };

        await activityProc.processIncomingActivity(turnContext);
    }

    private async sentCards (ctx: TurnContext) {
        let heroCard = CardFactory.heroCard(
            'Card title',
            undefined,
            CardFactory.actions([
                {
                    type: 'imBack',
                    title: 'imBack',
                    value: 'Test for imBack'
                },
                {
                    type: 'invoke',
                    title: 'invoke',
                    value: { key: 'invoke value' }
                }
            ]));

        let signinCard = CardFactory.signinCard('Signin', 'https://1355e2b4.ngrok.io/auth', 'Signin Card Test');

        let o365Card = teams.TeamsFactory.o365ConnectorCard({
            'summary': 'a o365 card',
            'themeColor': '#acd45f',
            'title': 'O365 card',
            'potentialAction': [
                <teams.O365ConnectorCardHttpPOST> {
                    '@type': 'HttpPOST',
                    '@id': 'justSubmit',
                    'name': 'Http POST',
                    'body': JSON.stringify({ key: 'value' })
                },
                <teams.O365ConnectorCardActionCard> {
                    '@type': 'ActionCard',
                    '@id': 'actionCard',
                    'name': 'Show Card',
                    'inputs': [
                        <teams.O365ConnectorCardTextInput> {
                            '@type': 'textInput',
                            'id': 'text-1',
                            'isMultiline': true
                        }
                    ],
                    'actions': [
                        <teams.O365ConnectorCardHttpPOST> {
                            '@type': 'HttpPOST',
                            '@id': 'submit',
                            'name': 'Http POST',
                            'body': JSON.stringify({
                                'text-1': '{{text-1.value}}'
                            })
                        }
                    ]
                }
            ]
        });

        const fileCard = teams.TeamsFactory.fileConsentCard(
            'icon-color.png', 
            {
                'description': 'This is the file I want to send you',
                'sizeInBytes': 3196,
                'acceptContext': {},
                'declineContext': {}
            }
        );

        await ctx.sendActivities([
            { attachments: [heroCard] },
            { attachments: [signinCard] },
            { attachments: [o365Card] },
            { attachments: [fileCard] }
        ]);
    }
}
