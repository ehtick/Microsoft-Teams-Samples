// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { App, IActivityContext } from '@microsoft/teams.apps';
import { Activity, TeamsChannelAccount, IMessageActivity } from '@microsoft/teams.api';

// Initialize Teams App - automatically uses CLIENT_ID and CLIENT_SECRET from environment variables
// Note: .env file is only required when running on Teams (not needed for local development with devtools)
const app = new App();

// Removes only the bot recipient mention from the incoming message text
// Note: This helper function should ideally be part of the SDK to avoid duplication across samples
function removeRecipientMention(activity: Activity): string {
    const messageActivity = activity as IMessageActivity;
    let text = messageActivity.text || '';
    
    if (!text || !activity.entities || !activity.recipient) {
        return text.trim();
    }

    // Only remove mentions that are specifically for the bot recipient
    for (const entity of activity.entities) {
        if (entity.type === 'mention' && 
            entity.mentioned?.id === activity.recipient.id && 
            entity.text) {
            // Replace only the first occurrence to avoid removing user mentions
            text = text.replace(entity.text, '').trim();
        }
    }
    
    return text;
}

// Handle conversation update events (when bot is added or members join)
app.on('conversationUpdate', async (context) => {
    const { activity } = context;
    const membersAdded = (activity as any).membersAdded || [];
    
    for (const member of membersAdded) {
        // Check if bot was added to the conversation
        if (member.id !== activity.recipient.id) {
            await sendWelcomeMessage(context);
        }
    }
});

// Handles incoming messages and routes to appropriate functions based on message content
app.on('message', async (context) => {
    const { activity } = context;
    console.log(`Received message: ${activity.text}`);

    // Remove bot mention from text
    let text = removeRecipientMention(activity);
    text = text.trim().toLowerCase();

    if (text.includes('mention me')) {
        await mentionUser(context);
    } else if (text.includes('who')) {
        await getSingleMember(context);
    } else {
        await sendWelcomeMessage(context);
    }
});

// Sends a welcome message with available commands
async function sendWelcomeMessage(context: IActivityContext): Promise<void> {
    await context.send({
        type: 'message',
        text: `Welcome to Microsoft Teams conversation update events demo bot.

Available commands:
- **mention me** - Bot will mention you in the reply
- **whoami** - Get your user information`
    });
}

// Retrieves and displays information about the current user
async function getSingleMember(context: IActivityContext): Promise<void> {
    const { activity } = context;
    const conversationId = activity.conversation.id;
    const userId = activity.from.id;
    
    try {
        const member: TeamsChannelAccount = await context.api.conversations.members(conversationId).getById(userId);
        await context.send({
            type: 'message',
            text: `You are: ${member.name}`
        });
    } catch (error) {
        console.error('Error getting member:', error);
        await context.send({
            type: 'message',
            text: 'Unable to retrieve member information. This feature may not be available in this context.'
        });
    }
}

// Mention a user in a message
async function mentionUser(context: IActivityContext): Promise<void> {
    const { activity } = context;
    const conversationId = activity.conversation.id;
    const userId = activity.from.id;
    
    try {
        const member: TeamsChannelAccount = await context.api.conversations.members(conversationId).getById(userId);
        
        // Create a text message with user mention
        const mentionText = `<at>${member.name}</at>`;
        await context.send({
            type: 'message',
            text: `Hello ${mentionText}`,
            entities: [
                {
                    type: 'mention',
                    text: mentionText,
                    mentioned: {
                        id: userId,
                        name: member.name,
                        role: 'user'
                    }
                }
            ]
        });
    } catch (error) {
        console.error('Error mentioning user:', error);
        await context.send({
            type: 'message',
            text: 'Unable to mention you. This feature may not be available in this context.'
        });
    }
}

// Starts the Teams bot application and listens for incoming requests
app.start().catch((err) => {
    console.error('Failed to start app:', err);
    process.exit(1);
});
