# This functionality is moving into the core Bot Framework SDK

We are migrating the functionality of this SDK into the core Bot Framework SDK, and are targeting the 4.6 release (early November 2019). Please see our [early example code](https://github.com/microsoft/botbuilder-dotnet/tree/master/tests/Teams) for an early look at the new, improved, way easier to use, SDK!

# Bot Builder SDK4 - Microsoft Teams Extensions

[![Build status](https://ci.appveyor.com/api/projects/status/4kghlfgmsw2tnk0o/branch/master?svg=true)](https://ci.appveyor.com/project/robin-liao/botbuilder-microsoftteams-node/branch/master)

The Microsoft Bot Builder SDK Teams Extensions allow bots built using Bot Builder SDK to consume Teams functionality easily. **[Review the documentation](https://msdn.microsoft.com/en-us/microsoft-teams/bots)** to get started!

# Samples

Get started quickly with our samples:

* Sample bots [here](https://github.com/OfficeDev/BotBuilder-MicrosoftTeams-node/tree/master/samples)


# This SDK allows you to easily...

* Fetch a list of channels in a team
* Fetch profile info about all members of a team
* Fetch tenant-id from an incoming message to bot
* Create 1:1 chat with a specific user
* Mention a specific user
* Consume various events like channel-created, team-renamed, etc.
* Accept messages only from specific tenants
* Write Compose Extensions
* _and more!_

# Getting started

* This SDK is the extension of [Bot Framework SDK 4](https://github.com/Microsoft/botbuilder-js), so you may start with [the quickstart](https://docs.microsoft.com/en-us/azure/bot-service/javascript/bot-builder-javascript-quickstart?view=azure-bot-service-4.0) or [the example for Azure Web app bot](https://docs.microsoft.com/en-us/azure/cognitive-services/luis/luis-nodejs-tutorial-bf-v4) if you never have experience for it. 

* Once you've got your bot scaffolded out, install the Teams BotBuilder package:

```bash
npm install botbuilder-teams@4.0.0-beta1
```

* To extend your bot to support Microsoft Teams, add middleware to adapter:
```typescript
// use Teams middleware
adapter.use(new teams.TeamsMiddleware());  
```
* Now in the `onTurn` method of your bot, to do any Teams specific stuff, first grab the `TeamsContext` as shown below:
```typescript
import { TeamsContext } from 'botbuilder-teams';

export class Bot {
  async onTurn(turnContext: TurnContext) {
    const teamsCtx: TeamsContext = TeamsContext.from(ctx);
  }
}
```
* And once you have `teamsContext`, you may utilize code autocomplete provided by Visual Studio Code to discover all the operations you can do. For instance, here's how you can fetch the list of channels in the team and fetch information about the team:
```typescript
// Now fetch the Team ID, Channel ID, and Tenant ID off of the incoming activity
const incomingTeamId = teamsCtx.team.id;
const incomingChannelid = teamsCtx.channel.id;
const incomingTenantId = teamsCtx.tenant.id;

// Make an operation call to fetch the list of channels in the team, and print count of channels.
var channels = await teamsCtx.teamsConnectorClient.teams.fetchChannelList(incomingTeamId);
await turnContext.sendActivity(`You have ${channels.conversations.length} channels in this team`);

// Make an operation call to fetch details of the team where the activity was posted, and print it.
var teamInfo = await teamsCtx.teamsConnectorClient.teams.fetchTeamDetails(incomingTeamId);
await turnContext.sendActivity(`Name of this team is ${teamInfo.name} and group-id is ${teamInfo.aadGroupId}`);
```

# Questions, bugs, feature requests, and contributions
Please review the information [here](https://msdn.microsoft.com/en-us/microsoft-teams/feedback).

# Contributing

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
