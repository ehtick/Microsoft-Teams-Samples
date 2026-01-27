# Bot Conversation - Python

This sample demonstrates a conversation bot for Microsoft Teams using Python and the Microsoft Teams SDK.

## Features

- Welcome cards with Hero Cards
- User mentions in Adaptive Cards
- Card updates with counter
- Card deletion
- Immersive Reader Card (Flight Status example)

## Prerequisites

- [Python 3.12+](https://www.python.org/downloads/)
- [uv](https://docs.astral.sh/uv/) (recommended) or pip

## Run the sample

1. Navigate to this directory:
   ```bash
   cd python
   ```

2. Install dependencies using uv:
   ```bash
   uv sync
   ```

3. Run the bot:
   ```bash
   uv run main.py
   ```

### Alternative: Using pip

```bash
pip install -e .
python main.py
```

The bot will start listening on `http://localhost:3978`.

## Run in the `agentsplayground`

Install the tool agentsplayground for your platform:

Windows:
```
winget install agentsplayground
```

Linux:
```
curl -s https://raw.githubusercontent.com/OfficeDev/microsoft-365-agents-toolkit/dev/.github/scripts/install-agentsplayground-linux.sh | bash
```

Other platforms (like MacOS, via npm):
```
npm install -g @microsoft/m365agentsplayground
```

Once the tool is installed, you can run it from your terminal with the command `agentsplayground`, and it will try to connect to `localhost:3978` where your bot is running.

## Run in the Teams Client

To run this sample in the Teams Client, you need to provision your app in a M365 Tenant, and configure the app to your DevTunnels URL.

1. Install the tool DevTunnels https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/get-started
2. Get Access to a M365 Developer Tenant https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-get-started
3. Create a Teams App with the Bot Feature in the Teams Developer Portal (in your tenant) https://dev.teams.microsoft.com

### Configure DevTunnels

Create a persistent tunnel for the port 3978 with anonymous access:

```
devtunnel create -a my-tunnel  
devtunnel port create -p 3978 my-tunnel  
devtunnel host my-tunnel
```

Take note of the URL shown after *Connect via browser:*

### Provisioning the Teams Application

Navigate to the Teams Developer Portal http://dev.teams.microsoft.com

#### Create a new Bot resource

1. Navigate to `Tools->Bot management`, and add a `New bot`
2. In Configure, paste the Endpoint address from devtunnels and append `/api/messages`
3. In Client secrets, create a new secret for later

#### Create a new Teams App

1. Navigate to `Apps` and create a `New App`
2. Fill the required values in Basic information (short and long name, short and long description and App URLs)
3. In `App features->Bot` select the bot you created previously
4. Select `Preview in Teams`

## Interacting with the bot

You can interact with this bot by sending it a message, or selecting a command from the command list. The bot will respond to the following strings:

1. **Who am I** - The bot will display information about the current user
2. **Find me in Adaptive Card** - The bot will send an Adaptive Card with user mentions (by UPN and AAD Object ID)
3. **Delete Card** - The bot will delete the card that triggered this action
4. **Update Card** - The bot will update the card with an incremented counter
5. **Send Immersive Reader Card** - The bot will send a flight status Adaptive Card with Immersive Reader support

## Further Reading

- [Microsoft Teams SDK Documentation](https://learn.microsoft.com/microsoftteams/platform/)
- [Bot Framework Documentation](https://docs.botframework.com)
