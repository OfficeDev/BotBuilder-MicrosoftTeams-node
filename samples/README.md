# Steps to run samples

Our samples are running along with the library version against **current repo** instead of NPM release library, so you may preview the sample code for several pre-release features. (So that's why you'll see `"botbuilder-teams": "file:../../botbuilder-teams-js"` in `package.json` in sample folders)

1. **Build library:** go to lib path `~/BotBuilder-MicrosoftTeams-node/botbuilder-teams-js`:
    1. install and build: `npm i`
    2. for post-install and you just wanna build or re-build: `npm run build`

2. **Set up credentials**: in sample folders find out bot file (usually named `bot-file.json`) and set up your Microsoft Bot ID (App ID) and associated password for `appId` and `appPassword` fields respectively.

3. **Run up samples**: then follow the steps to run up bot:
    1. `npm i`
    2. `npm start`

4. **Use bot in Teams as sideloaded app**: to use the bot as a sideloaded app in Microsoft Teams, please follow the steps:
    1. Modify `manifest.json` to assign a random GUID for `id` as Microsoft Teams App ID
    2. Assign `bots.botId` and `composeExtensions.botId` where the bot ID should be the one assigned in step 2 in `bot-file.json`
    3. Zip `manifest.json`, `icon-color.png` and `icon-outline.png` as an archived file. 
    4. Follow [instructions](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload) to side-load your bot into Teams.
    5. More instruction can be found [here](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-nodejs-app-studio)