// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { App, IActivityContext } from '@microsoft/teams.apps';
import { TeamsChannelAccount, IMessageActivity } from '@microsoft/teams.api';

// Initialize Teams App - automatically uses CLIENT_ID and CLIENT_SECRET from environment variables
// Note: .env file is only required when running on Teams (not needed for local development with devtools)
const app = new App();

// Simple in-memory storage for conversation context (for proactive messaging)
// In production, use persistent storage like a database
const contextStorage: Map<string, IActivityContext> = new Map();

// Send a proactive message to a user
async function sendProactiveNotification(
    userId: string,
    message: string = 'Hey! This is a proactive message from the bot!'
): Promise<boolean> {
    const storedContext = contextStorage.get(userId);
    if (!storedContext) {
        return false;
    }

    try {
        // Reuse the stored context to send the message
        await storedContext.send({
            type: 'message',
            text: message
        });
        return true;
    } catch (error) {
        console.error('Failed to send proactive message:', error);
        return false;
    }
}

// Send a proactive message after a delay
async function delayedProactiveMessage(userId: string, delaySeconds: number = 10): Promise<void> {
    setTimeout(async () => {
        try {
            await sendProactiveNotification(
                userId,
                `Reminder: This proactive message was sent ${delaySeconds} seconds after your request!`
            );
        } catch (err) {
            console.error('Error sending proactive message:', err);
        }
    }, delaySeconds * 1000);
}

// Handle conversation update events (when bot is added or members join)
app.on('conversationUpdate', async (context) => {
    const { activity } = context;
    const membersAdded = (activity as any).membersAdded || [];

    for (const member of membersAdded) {
        // Check if bot was added to the conversation
        if (member.id === activity.recipient.id) {
            await sendWelcomeMessage(context);
        }
    }
});

// Handles incoming messages and routes to appropriate functions based on message content
app.on('message', async (context) => {
    const { activity } = context;

    // Get message text and normalize it
    const messageActivity = activity as IMessageActivity;
    let text = (messageActivity.text || '').trim().toLowerCase();

    // Store conversation context for proactive messaging (from any message)
    const userAadId = (activity.from as any).aadObjectId;
    if (userAadId) {
        contextStorage.set(userAadId, context);
    }

    // Handle proactive messaging command
    if (text.includes('proactive')) {
        if (userAadId) {
            await context.send({
                type: 'message',
                text: "Got it! I'll send you a proactive message in 10 seconds..."
            });
            // Schedule the proactive message (runs in background)
            delayedProactiveMessage(userAadId, 10);
        } else {
            await context.send({
                type: 'message',
                text: "Sorry, I couldn't identify your user ID for proactive messaging."
            });
        }
        return;
    }

    // Handle mention me command
    if (text.includes('mentionme') || text.includes('mention me')) {
        await mentionUser(context);
    }
    // Handle whoami command
    else if (text.includes('whoami')) {
        await getSingleMember(context);
    }
    // Handle welcome command
    else if (text.includes('welcome')) {
        await sendWelcomeMessage(context);
    }
    // Handle greeting messages
    else if (text.includes('hi') || text.includes('hello')) {
        await echoMessage(context, text);
    }
    // Default: echo back any other message
    else if (text) {
        await echoMessage(context, text);
    }
});

// Sends a welcome message
async function sendWelcomeMessage(context: IActivityContext): Promise<void> {
    await context.send({
        type: 'message',
        text: 'Welcome to the Teams Quickstart Bot!'
    });
}

// Echo back the user's message
async function echoMessage(context: IActivityContext, text: string): Promise<void> {
    await context.send({
        type: 'message',
        text: `**Echo:** ${text}`
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
    }
}

// Starts the Teams bot application and listens for incoming requests
app.start().catch(console.error);
