# Overview of NorthwindDB-MSTeamExt-Sample

This app showcase how to access a AAD secured REST API with SSO from a M365 Copilot plugin developed by using ai-library. 

## Get started with the template
* this sameple was original created based on M365 copilot northwind plugin sample. the original sample used Azure storage DB. The sample introduced a REST API to access Northwind DB from a Cosmos DB. For the copmlete REST API sample, please refer to [frankchen76 / northwinddb-api](https://github.com/frankchen76/northwinddb-api)
* follow [Add single sign-on to Teams app](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/add-single-sign-on?tabs=typescript%2F%3Ffrom%3Dteamstoolkit&pivots=visual-studio-code-v5) to add SSO which will 
* create .env.local under env folder. 
```
# This file includes environment variables that can be committed to git. It's gitignored by default because it represents your local development environment.

# Built-in environment variables
TEAMSFX_ENV=local
APP_NAME_SUFFIX=local

# Generated during provision, you can also add your own variables.
BOT_ID=
TEAMS_APP_ID=
BOT_DOMAIN=
BOT_ENDPOINT=
# NorthwindDb API
NORTHWINDDBAPI_ENDPOINT=[rest-api-endpoint]
# Debug flag
DEBUG=northwinddb-msteamext
```
* create .env.local.user under env folder. 
```
# This file includes environment variables that will not be committed to git by default. You can set these environment variables in your CI/CD system for your project.

# If you're adding a secret value, add SECRET_ prefix to the name so Teams Toolkit can handle them properly
# Secrets. Keys prefixed with `SECRET_` will be masked in Teams Toolkit logs.
SECRET_BOT_PASSWORD=
```
* the smaple uses ngrok with predefine domain, if you want to use OOTB dev tunnel, please uncomment out "Start local tunnel" in .vscode/tasks.json. like below
```
{
            "label": "Start Teams App Locally",
            "dependsOn": [
                "Validate prerequisites",
                "Start local tunnel",
                "Provision",
                "Deploy",
                "Start application"
            ],
            "dependsOrder": "sequence"
        },
```

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 16, 18
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
> - [Set up your dev environment for extending Teams apps across Microsoft 365](https://aka.ms/teamsfx-m365-apps-prerequisites)
> Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect.
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teamsfx-cli)

1. First, select the Teams Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug (Edge)` or `Debug (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
5. To trigger the Message Extension, you can:
   1. In Teams: `@mention` Your message extension from the `search box area`, `@mention` your message extension from the `compose message area` or click the `...` under compose message area to find your message extension.
   2. In Outlook: click the `More apps` icon under compose email area to find your message extension.

