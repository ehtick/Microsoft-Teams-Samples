# Bot Quickstart - TypeScript

This sample demonstrates a basic Teams Conversation Bot for Microsoft Teams using TypeScript with conversation features including message handling and user mentions.

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)

## Run the sample

1. Navigate to this directory:
   ```
   cd nodejs/bot-quickstart
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Run the bot:
   ```
   npm start
   ```

The bot will start listening on `http://localhost:3978`.

## Features

- **Conversation Update Events**: Sends welcome message when bot is added or user joins
- **Message Handling**: Responds to user messages with basic conversation features
- **User Information**: Get current user details with "whoami" command
- **User Mentions**: Mention users in messages with "mention me" command

## Configuration

**Note:** The `.env` file is only required when running on Teams (not needed for local development).

Create a `.env` file (use `.env.TEMPLATE` as reference) with your credentials:

```
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
```

Refer to the main [README.md](../../README.md) to interact with your bot in the agentsplayground or in Teams.
