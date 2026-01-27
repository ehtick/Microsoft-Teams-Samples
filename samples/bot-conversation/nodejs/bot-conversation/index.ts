// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { App } from '@microsoft/teams.apps';
import * as ACData from 'adaptivecards-templating';

// Initialize Teams App - automatically uses CLIENT_ID and CLIENT_SECRET from environment variables
const app = new App();

// Removes bot mention text from the incoming message
function removeMentionText(text: string, entities: any[]): string {
    if (!text || !entities) return text || '';

    for (const entity of entities) {
        if (entity.type === 'mention' && entity.text) {
            text = text.replace(entity.text, '').trim();
        }
    }
    return text;
}

// Handles incoming messages and routes to appropriate functions based on message content
app.on('message', async (context) => {
    const { activity } = context;

    // Remove bot mention from text
    let text = removeMentionText(activity.text || '', activity.entities || []);
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
async function cardActivity(context: any, isUpdate: boolean) {
    const cardActions = [
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
async function sendUpdateCard(context: any, cardActions: any[]) {
    const { activity } = context;
    const data = activity.value || { count: 0 };
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
async function sendWelcomeCard(context: any, cardActions: any[]) {
    const initialValue = { count: 0 };

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
async function getSingleMember(context: any) {
    try {
        const { activity } = context;
        const conversationId = activity.conversation.id;
        const userId = activity.from.id;

        const member = await context.api.conversations.members(conversationId).getById(userId);

        await context.send({
            type: 'message',
            text: `You are: ${member.name}`
        });
    } catch (e: any) {
        console.error('Error getting member:', e);
        await context.send({
            type: 'message',
            text: 'Member not found or error occurred.'
        });
    }
}

// Adaptive card template for mentioning users by UPN and AAD Object ID
const UserMentionCardTemplate = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.5",
    "speak": "This card mentions a user by User Principle Name: Hello ${userName}",
    "body": [
        {
            "type": "TextBlock",
            "text": "Mention a user by User Principle Name: Hello <at>${userName} UPN</at>"
        },
        {
            "type": "TextBlock",
            "text": "Mention a user by AAD Object Id: Hello <at>${userName} AAD</at>"
        }
    ],
    "msteams": {
        "entities": [
            {
                "type": "mention",
                "text": "<at>${userName} UPN</at>",
                "mentioned": {
                    "id": "${userUPN}",
                    "name": "${userName}"
                }
            },
            {
                "type": "mention",
                "text": "<at>${userName} AAD</at>",
                "mentioned": {
                    "id": "${userAAD}",
                    "name": "${userName}"
                }
            }
        ]
    }
};

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
async function mentionAdaptiveCardActivity(context: any) {
    try {
        const { activity } = context;
        const conversationId = activity.conversation.id;
        const userId = activity.from.id;

        const member = await context.api.conversations.members(conversationId).getById(userId);

        const template = new ACData.Template(UserMentionCardTemplate);
        const memberData = {
            userName: member.name,
            userUPN: member.userPrincipalName,
            userAAD: member.aadObjectId
        };

        const adaptiveCard = template.expand({
            $root: memberData
        });

        await context.send({
            type: 'message',
            attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: adaptiveCard
            }]
        });
    } catch (e: any) {
        console.error('Error getting member:', e);
        await context.send({
            type: 'message',
            text: 'Member not found or error occurred.'
        });
    }
}

// Sends an adaptive card with Immersive Reader support showing flight information
async function getImmersivereaderCard(context: any) {
    await context.send({
        type: 'message',
        attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: ImmersiveReaderCardTemplate
        }]
    });
}

// Deletes a card message from the conversation
async function deleteCardActivity(context: any) {
    const { activity } = context;
    const conversationId = activity.conversation.id;
    const messageId = activity.replyToId;

    await context.api.conversations.activities(conversationId).delete(messageId);
}

// Starts the Teams bot application and listens for incoming requests
app.start().catch((err) => {
    console.error('Failed to start app:', err);
    process.exit(1);
});
