# Bot Conversation - Node.js (TypeScript)

This sample demonstrates a comprehensive Teams Conversation Bot for Microsoft Teams using Node.js and TypeScript. The bot showcases various conversation features including user mentions, adaptive cards, card updates, and immersive reader support.

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS version recommended)

## Run the sample

1. Navigate to this directory:
   ```
   cd nodejs/bot-conversation
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

- **Welcome Card**: Interactive hero card with multiple action buttons
- **User Information**: Get current user details with "Who am I?" functionality
- **Adaptive Card Mentions**: Mention users in adaptive cards using UPN and AAD Object ID
- **Card Management**: Update and delete hero cards dynamically
- **Immersive Reader**: Send adaptive cards with Immersive Reader support for accessibility

## Configuration

`.env` file with your credentials:

```
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
```

Refer to the main [README.md](../../README.md) to interact with your bot in the agentsplayground or in Teams.
