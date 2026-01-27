// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { App, IActivityContext } from '@microsoft/teams.apps';
import { Activity, TeamsChannelAccount, IMessageActivity } from '@microsoft/teams.api';

// Type definitions for card actions
interface CardAction {
    type: string;
    title: string;
    value: CardUpdateData | null;
    text: string;
}

interface CardUpdateData {
    count: number;
}

// Initialize Teams App - automatically uses CLIENT_ID and CLIENT_SECRET from environment variables
const app = new App();

// Removes only the bot recipient mention from the incoming message text
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

// Handles incoming messages and routes to appropriate functions based on message content
app.on('message', async (context) => {
    const { activity } = context;
    console.log(`Received message: ${activity.text}`);

    // Remove bot mention from text
    let text = removeRecipientMention(activity);
    text = text.trim().toLowerCase();

    if (text.includes('mention me')) {
        await mentionAdaptiveCardActivity(context);
    } else if (text.includes('update')) {
        await cardActivity(context, true);
    } else if (text.includes('delete')) {
        await deleteCardActivity(context);
    } else if (text.includes('who')) {
        await getSingleMember(context);
    } else if (text.includes('immersivereader')) {
        await getImmersivereaderCard(context);
    } else {
        await cardActivity(context, false);
    }
});

// Creates and sends either a welcome card or update card based on isUpdate flag
async function cardActivity(context: IActivityContext, isUpdate: boolean): Promise<void> {
    const cardActions: CardAction[] = [
        {
            type: 'messageBack',
            title: 'Who am I?',
            value: null,
            text: 'whoami'
        },
        {
            type: 'messageBack',
            title: 'Find me in Adaptive Card',
            value: null,
            text: 'mention me'
        },
        {
            type: 'messageBack',
            title: 'Delete card',
            value: null,
            text: 'Delete'
        },
        {
            type: 'messageBack',
            title: 'Send Immersive Reader Card',
            value: null,
            text: 'ImmersiveReader'
        }
    ];

    if (isUpdate) {
        await sendUpdateCard(context, cardActions);
    } else {
        await sendWelcomeCard(context, cardActions);
    }
}

// Updates an existing card with incremented count
async function sendUpdateCard(context: IActivityContext, cardActions: CardAction[]): Promise<void> {
    const { activity } = context;
    const messageActivity = activity as IMessageActivity;
    const data: CardUpdateData = messageActivity.value || { count: 0 };
    data.count += 1;

    cardActions.push({
        type: 'messageBack',
        title: 'Update Card',
        value: data,
        text: 'UpdateCardAction'
    });

    // Update the card using the API
    const conversationId = activity.conversation.id;
    const messageId = activity.replyToId;
    
    if (!messageId) return;

    await context.api.conversations.activities(conversationId).update(messageId, {
        type: 'message',
        attachments: [{
            contentType: 'application/vnd.microsoft.card.hero',
            content: {
                title: 'Updated card',
                text: `Update count: ${data.count}`,
                buttons: cardActions
            }
        }]
    });
}

// Sends initial welcome card with action buttons
async function sendWelcomeCard(context: IActivityContext, cardActions: CardAction[]): Promise<void> {
    const initialValue: CardUpdateData = { count: 0 };

    cardActions.push({
        type: 'messageBack',
        title: 'Update Card',
        value: initialValue,
        text: 'UpdateCardAction'
    });

    await context.send({
        type: 'message',
        attachments: [{
            contentType: 'application/vnd.microsoft.card.hero',
            content: {
                title: 'Welcome card',
                text: '',
                buttons: cardActions
            }
        }]
    });
}

// Retrieves and displays information about the current user
async function getSingleMember(context: IActivityContext): Promise<void> {
    try {
        const { activity } = context;
        const conversationId = activity.conversation.id;
        const userId = activity.from.id;
        
        const member: TeamsChannelAccount = await context.api.conversations.members(conversationId).getById(userId);
        await context.send({
            type: 'message',
            text: `You are: ${member.name}`
        });
    } catch (error) {
        console.error('Error getting member:', error);
        await context.send({
            type: 'message',
            text: 'Member not found or error occurred.'
        });
    }
}

// Adaptive card template for Immersive Reader with flight information example
const ImmersiveReaderCardTemplate = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.5",
    "speak": "Flight KL0605 to San Fransisco has been delayed.It will not leave until 10:10 AM.",
    "body": [
        {
            "type": "ColumnSet",
            "columns": [
                {
                    "type": "Column",
                    "width": "auto",
                    "items": [
                        {
                            "type": "Image",
                            "size": "Small",
                            "url": "https://adaptivecards.io/content/airplane.png",
                            "altText": "Airplane"
                        }
                    ]
                },
                {
                    "type": "Column",
                    "width": "stretch",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Flight Status",
                            "horizontalAlignment": "Right",
                            "isSubtle": true,
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "DELAYED",
                            "horizontalAlignment": "Right",
                            "spacing": "None",
                            "size": "Large",
                            "color": "Attention",
                            "wrap": true
                        }
                    ]
                }
            ]
        },
        {
            "type": "ColumnSet",
            "separator": true,
            "spacing": "Medium",
            "columns": [
                {
                    "type": "Column",
                    "width": "stretch",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Passengers",
                            "isSubtle": true,
                            "weight": "Bolder",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "Sarah Hum",
                            "spacing": "Small",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "Jeremy Goldberg",
                            "spacing": "Small",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "Evan Litvak",
                            "spacing": "Small",
                            "wrap": true
                        }
                    ]
                },
                {
                    "type": "Column",
                    "width": "auto",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Seat",
                            "horizontalAlignment": "Right",
                            "isSubtle": true,
                            "weight": "Bolder",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "14A",
                            "horizontalAlignment": "Right",
                            "spacing": "Small",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "14B",
                            "horizontalAlignment": "Right",
                            "spacing": "Small",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "14C",
                            "horizontalAlignment": "Right",
                            "spacing": "Small",
                            "wrap": true
                        }
                    ]
                }
            ]
        },
        {
            "type": "ColumnSet",
            "spacing": "Medium",
            "separator": true,
            "columns": [
                {
                    "type": "Column",
                    "width": 1,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Flight",
                            "isSubtle": true,
                            "weight": "Bolder",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "KL605",
                            "spacing": "Small",
                            "wrap": true
                        }
                    ]
                },
                {
                    "type": "Column",
                    "width": 1,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Departs",
                            "isSubtle": true,
                            "horizontalAlignment": "Center",
                            "weight": "Bolder",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "{{TIME(2017-03-04T09:20:00-01:00)}}",
                            "color": "Attention",
                            "weight": "Bolder",
                            "horizontalAlignment": "Center",
                            "spacing": "Small",
                            "wrap": true
                        }
                    ]
                },
                {
                    "type": "Column",
                    "width": 1,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Arrives",
                            "isSubtle": true,
                            "horizontalAlignment": "Right",
                            "weight": "Bolder",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "{{TIME(2017-03-05T08:20:00+04:00)}}",
                            "color": "Attention",
                            "horizontalAlignment": "Right",
                            "weight": "Bolder",
                            "spacing": "Small",
                            "wrap": true
                        }
                    ]
                }
            ]
        },
        {
            "type": "ColumnSet",
            "spacing": "Medium",
            "separator": true,
            "columns": [
                {
                    "type": "Column",
                    "width": 1,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Amsterdam Airport",
                            "isSubtle": true,
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "AMS",
                            "size": "ExtraLarge",
                            "color": "Accent",
                            "spacing": "None",
                            "wrap": true
                        }
                    ]
                },
                {
                    "type": "Column",
                    "width": "auto",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": " ",
                            "wrap": true
                        },
                        {
                            "type": "Image",
                            "url": "https://adaptivecards.io/content/airplane.png",
                            "altText": "Airplane",
                            "size": "Small"
                        }
                    ]
                },
                {
                    "type": "Column",
                    "width": 1,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "San Francisco Airport",
                            "isSubtle": true,
                            "horizontalAlignment": "Right",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "SFO",
                            "horizontalAlignment": "Right",
                            "size": "ExtraLarge",
                            "color": "Accent",
                            "spacing": "None",
                            "wrap": true
                        }
                    ]
                }
            ]
        }
    ]
};

// Sends an adaptive card that mentions the user by UPN and AAD Object ID
async function mentionAdaptiveCardActivity(context: IActivityContext): Promise<void> {
    try {
        const { activity } = context;
        const conversationId = activity.conversation.id;
        const userId = activity.from.id;
        
        const member: TeamsChannelAccount = await context.api.conversations.members(conversationId).getById(userId);
        
        // Use aadObjectId if objectId is not available, or fall back to userId
        const aadId = (member as any).aadObjectId || member.objectId || userId;
        
        if (!member.userPrincipalName) {
            console.error('Member UPN is missing');
            await context.send({
                type: 'message',
                text: 'Unable to create mention card: user information incomplete.'
            });
            return;
        }
        
        // Manually build the adaptive card with proper mentions
        const adaptiveCard = {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "speak": `This card mentions a user by User Principle Name: Hello ${member.name}`,
            "body": [
                {
                    "type": "TextBlock",
                    "text": `Mention a user by User Principle Name: Hello <at>${member.name}</at>`
                },
                {
                    "type": "TextBlock",
                    "text": `Mention a user by AAD Object Id: Hello <at>${member.name}</at>`
                }
            ],
            "msteams": {
                "entities": [
                    {
                        "type": "mention",
                        "text": `<at>${member.name}</at>`,
                        "mentioned": {
                            "id": member.userPrincipalName,
                            "name": member.name
                        }
                    },
                    {
                        "type": "mention",
                        "text": `<at>${member.name}</at>`,
                        "mentioned": {
                            "id": aadId,
                            "name": member.name
                        }
                    }
                ]
            }
        };

        await context.send({
            type: 'message',
            attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: adaptiveCard
            }]
        });
    } catch (error) {
        console.error('Error getting member:', error);
        await context.send({
            type: 'message',
            text: 'Member not found or error occurred.'
        });
    }
}

// Sends an adaptive card with Immersive Reader support showing flight information
async function getImmersivereaderCard(context: IActivityContext): Promise<void> {
    await context.send({
        type: 'message',
        attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: ImmersiveReaderCardTemplate
        }]
    });
}

// Deletes a card message from the conversation
async function deleteCardActivity(context: IActivityContext): Promise<void> {
    const { activity } = context;
    const conversationId = activity.conversation.id;
    const messageId = activity.replyToId;

    if (!messageId) return;
    
    await context.api.conversations.activities(conversationId).delete(messageId);
}

// Starts the Teams bot application and listens for incoming requests
app.start().catch((err) => {
    console.error('Failed to start app:', err);
    process.exit(1);
});
