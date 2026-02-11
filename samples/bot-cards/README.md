# Cards and Attachments Bot

This sample demonstrates various card types and attachment handling in Microsoft Teams using a bot built with Teams SDK. The bot showcases three key functionalities: **Adaptive Card Actions** for dynamic card behaviors, **Toggle Visibility Cards** for showing/hiding content, and **File Upload** capabilities for managing attachments and inline images. This comprehensive sample illustrates how to create rich, interactive bot experiences in Teams using cards and file management features.

## Table of Contents

- [Interaction with Bot](#interaction-with-bot)
- [Sample Implementations](#sample-implementations)
- [How to run these samples](#how-to-run-these-samples)
  - [Run in the agentsplayground](#run-in-the-agentsplayground)
  - [Run in the Teams Client](#run-in-the-teams-client)
    - [Configure DevTunnels](#configure-devtunnels)
    - [Provisioning the Teams Application](#provisioning-the-teams-application)
  - [Configure the new project to use the new Teams Bot Application](#configure-the-new-project-to-use-the-new-teams-bot-application)
  - [Pro Tip: Read the configuration settings using the Azure CLI](#pro-tip-read-the-configuration-settings-using-the-azure-cli)
- [Troubleshooting](#troubleshooting)
- [Further Reading](#further-reading)

## Interaction with Bot

![File Upload](bot_cards.gif)

The bot supports the following functionalities:

### 1. Adaptive Card Actions
* Interactive cards with submit actions and button interactions
* Dynamic card behaviors based on user input
* Demonstrates various adaptive card action types

### 2. Toggle Visibility Card
* Shows/hides content dynamically within cards
* Demonstrates collapsible sections in adaptive cards
* Provides better content organization and user experience

### 3. File Upload
* **Accept file upload** - Uploads the `teams-logo.png` file from the Files directory
* **Decline file upload** - Cancels the file upload operation
* **Send file attachment** - Bot receives and saves the file sent as an attachment
* **Send inline image** - Bot processes inline images from the message compose section

## Sample Implementations

| Language | Framework | Directory |
|----------|-----------|-----------|
| C# | .NET 10 / ASP.NET Core | [dotnet](dotnet/bot-cards/README.md) |
| Typescript | Node.js | [nodejs](nodejs/bot-cards/README.md) |
| Python | Python | [python](python/bot-cards/README.md) |

# How to run these samples

You can run these samples locally using

1. The agentsplayground tool, without provisioning the Teams App, or
2. In the Teams Client after you have provisioned the Teams Application and configured the application with your local DevTunnels URL.

## Run in the `agentsplayground`

Install the tool agentsplayground for your platform

Windows

```
winget install agentsplayground
```

Linux

```
curl -s https://raw.githubusercontent.com/OfficeDev/microsoft-365-agents-toolkit/dev/.github/scripts/install-agentsplayground-linux.sh | bash
```

Other platforms (like MacOS, via npm)

```
npm install -g @microsoft/m365agentsplayground
```

Once the tool is installed, you can run it from your terminal with the command `agentsplayground`, and it will try to connect to `localhost:3978` where your bot is running.

## Run in the Teams Client

To run these samples in the Teams Client, you need to provision your app in a M365 Tenant, and configure the app to your DevTunnels URL.

1. Install the tool DevTunnels https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/get-started
2. Get Access to a M365 Developer Tenant https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-get-started
3. Create a Teams App with the Bot Feature in the Teams Developer Portal (in your tenant) https://dev.teams.microsoft.com

### Configure DevTunnels

Create a persistent tunnel for the port 3978 with anonymous access

```
devtunnel create -a my-tunnel  
devtunnel port create -p 3978  my-tunnel 
devtunnel host  my-tunnel
```

Take note of the URL shown after *Connect via browser:*

### Provisioning the Teams Application

Navigate to the Teams Developer Portal http://dev.teams.microsoft.com

#### Create a new Bot resource

1. Navigate to `Tools->Bot management`, and add a `New bot`
1. In Configure, paste the Endpoint address from devtunnels and append `/api/messages`
1. In Client secrets, create a new secret and save it for later

> Note. If you have access to an Azure Subscription in the same Tenant, you can also create the Azure Bot resource ([learn more](https://learn.microsoft.com/en-us/azure/bot-service/abs-quickstart?view=azure-bot-service-4.0&tabs=singletenant)).

#### Create a new Teams App

1. Navigate to `Apps` and create a `New App`
1. Fill the required values in Basic information (short and long name, short and long description and App URLs)
1. In `App features->Bot` select the bot you created previously
1. Select `Preview in Teams`

> Note. When using an Azure Bot resource, provide the ClientID instead of selecting an existing bot.

## Configure the new project to use the new Teams Bot Application

For NodeJS and Python you will need a `.env` file with the next fields

```
TENANT_ID=
CLIENT_ID=
CLIENT_SECRET=
```

For dotnet you need to add these values to `appsettings.json` or `launchSettings.json` using the next syntax.

appSettings.json

```json
"urls" : "http://localhost:3978",
"Teams": {
    "ClientID": "",
    "ClientSecret": "",
    "TenantId": ""
  },
```

Or to use Env Vars from the profile defined in `launchSettings.json` (using the Environment Configuration Provider)

```json
"teamsbot": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": false,
      "applicationUrl": "http://localhost:3978",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development",
        "Teams__TenantId": "YOUR_TenantId",
        "Teams__ClientID": "YOUR_ClientId",
        "Teams__ClientSecret": "YOUR_ClientSecret"
      }
    }
```

## Pro Tip: Read the configuration settings using the Azure CLI

To obtain the TenantId, ClientId and SecretId you can use the Azure CLI with:

> Note. If you don't have access to an Azure Subscription you can still use the Azure CLI, make sure you login with `az login --allow-no-subscription`

```
az ad app credential reset --id $appId
```

## Troubleshooting

- If Teams cannot communicate with your bot, verify your DevTunnels URL is reachable.
- Ensure your .env or appsettings file is setup correctly.
- For file upload issues, verify the bot has permission to receive file uploads and the Files directory exists with proper write permissions.

## Further Reading

- [Upload Files Using Bots](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/bots-filesv4)
- [Microsoft Teams SDK Documentation](https://learn.microsoft.com/microsoftteams/platform/)
