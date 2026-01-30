"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

import asyncio
import json
import re
from typing import Dict, List, Optional
import copy

from dotenv import load_dotenv
from microsoft_teams.api import MessageActivity, TypingActivityInput, MessageActivityInput
from microsoft_teams.apps import ActivityContext, App

# Load environment variables
load_dotenv()

# Adaptive Card template for user mentions
USER_MENTION_CARD_TEMPLATE = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [
        {
            "type": "TextBlock",
            "text": "Mention a user by User Principal Name: Hello <at>${userName} UPN</at>"
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
}

# Immersive Reader Card template (Flight Status)
IMMERSIVE_READER_CARD_TEMPLATE = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.5",
    "speak": "Flight KL0605 to San Francisco has been delayed. It will not leave until 10:10 AM.",
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
                            "isSubtle": True,
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "DELAYED",
                            "horizontalAlignment": "Right",
                            "spacing": "None",
                            "size": "Large",
                            "color": "Attention",
                            "wrap": True
                        }
                    ]
                }
            ]
        },
        {
            "type": "ColumnSet",
            "separator": True,
            "spacing": "Medium",
            "columns": [
                {
                    "type": "Column",
                    "width": "stretch",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Passengers",
                            "isSubtle": True,
                            "weight": "Bolder",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Sarah Hum",
                            "spacing": "Small",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Jeremy Goldberg",
                            "spacing": "Small",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Evan Litvak",
                            "spacing": "Small",
                            "wrap": True
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
                            "isSubtle": True,
                            "weight": "Bolder",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "14A",
                            "horizontalAlignment": "Right",
                            "spacing": "Small",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "14B",
                            "horizontalAlignment": "Right",
                            "spacing": "Small",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "14C",
                            "horizontalAlignment": "Right",
                            "spacing": "Small",
                            "wrap": True
                        }
                    ]
                }
            ]
        },
        {
            "type": "ColumnSet",
            "spacing": "Medium",
            "separator": True,
            "columns": [
                {
                    "type": "Column",
                    "width": 1,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Flight",
                            "isSubtle": True,
                            "weight": "Bolder",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "KL605",
                            "spacing": "Small",
                            "wrap": True
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
                            "isSubtle": True,
                            "horizontalAlignment": "Center",
                            "weight": "Bolder",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "{{TIME(2017-03-04T09:20:00-01:00)}}",
                            "color": "Attention",
                            "weight": "Bolder",
                            "horizontalAlignment": "Center",
                            "spacing": "Small",
                            "wrap": True
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
                            "isSubtle": True,
                            "horizontalAlignment": "Right",
                            "weight": "Bolder",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "{{TIME(2017-03-05T08:20:00+04:00)}}",
                            "color": "Attention",
                            "horizontalAlignment": "Right",
                            "weight": "Bolder",
                            "spacing": "Small",
                            "wrap": True
                        }
                    ]
                }
            ]
        },
        {
            "type": "ColumnSet",
            "spacing": "Medium",
            "separator": True,
            "columns": [
                {
                    "type": "Column",
                    "width": 1,
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Amsterdam Airport",
                            "isSubtle": True,
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "AMS",
                            "size": "ExtraLarge",
                            "color": "Accent",
                            "spacing": "None",
                            "wrap": True
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
                            "wrap": True
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
                            "isSubtle": True,
                            "horizontalAlignment": "Right",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "SFO",
                            "horizontalAlignment": "Right",
                            "size": "ExtraLarge",
                            "color": "Accent",
                            "spacing": "None",
                            "wrap": True
                        }
                    ]
                }
            ]
        }
    ]
}

app = App()


def create_hero_card(title: str, text: str, buttons: List[Dict]) -> Dict:
    """Create a Hero Card structure."""
    card = {
        "contentType": "application/vnd.microsoft.card.hero",
        "content": {
            "title": title,
            "text": text,
            "buttons": buttons
        }
    }
    return card


def get_welcome_card_buttons(update_value: Optional[Dict] = None) -> List[Dict]:
    """Get the buttons for the welcome/update card."""
    buttons = [
        {
            "type": "messageBack",
            "title": "Who am I?",
            "text": "whoami"
        },
        {
            "type": "messageBack",
            "title": "Find me in Adaptive Card",
            "text": "mention me"
        },
        {
            "type": "messageBack",
            "title": "Delete card",
            "text": "deletecard"
        },
        {
            "type": "messageBack",
            "title": "Send Immersive Reader Card",
            "text": "immersivereader"
        }
    ]
    
    # Add update card button
    update_button = {
        "type": "messageBack",
        "title": "Update Card",
        "text": "updatecardaction",
        "value": json.dumps(update_value if update_value else {"count": 0})
    }
    buttons.append(update_button)
    
    return buttons


async def send_welcome_card(ctx: ActivityContext[MessageActivity]) -> None:
    """Send the welcome Hero Card."""
    buttons = get_welcome_card_buttons()
    card = create_hero_card(
        title="Welcome Card",
        text="Click the buttons.",
        buttons=buttons
    )
    
    activity = MessageActivityInput(
        text="",
        attachments=[card]
    )
    await ctx.send(activity)


async def send_update_card(ctx: ActivityContext[MessageActivity]) -> None:
    """Send an updated Hero Card with incremented count."""
    # Get the value from the activity
    value = ctx.activity.value if hasattr(ctx.activity, 'value') and ctx.activity.value else {}
    
    # Parse value if it's a string
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            value = {}
    
    # Increment count
    count = value.get("count", 0) + 1
    new_value = {"count": count}
    
    buttons = get_welcome_card_buttons(new_value)
    
    card = create_hero_card(
        title="Updated card",
        text=f"Update count {count}",
        buttons=buttons
    )
    
    # Create the updated activity
    activity = MessageActivityInput(
        text="",
        attachments=[card]
    )
    
    # Update the existing activity
    if hasattr(ctx.activity, 'reply_to_id') and ctx.activity.reply_to_id:
        await ctx.api.conversations.activities(
            ctx.activity.conversation.id
        ).update(ctx.activity.reply_to_id, activity)
    else:
        await ctx.send(activity)


async def mention_adaptive_card_activity(ctx: ActivityContext[MessageActivity]) -> None:
    """Send an Adaptive Card with user mentions."""
    try:
        # Get member information
        conversation_id = ctx.activity.conversation.id
        user_id = ctx.activity.from_.id
        
        # Try to get member details from the API
        try:
            members = await ctx.api.conversations.members(conversation_id).get_all()
            member = None
            for m in members:
                if m.id == user_id:
                    member = m
                    break
            
            if not member:
                member = ctx.activity.from_
        except Exception:
            # Fallback to activity sender info
            member = ctx.activity.from_
        
        member_name = member.name if hasattr(member, 'name') and member.name else "User"
        member_upn = member.user_principal_name if hasattr(member, 'user_principal_name') and member.user_principal_name else member.id
        member_aad = member.aad_object_id if hasattr(member, 'aad_object_id') and member.aad_object_id else member.id
        
        template = copy.deepcopy(USER_MENTION_CARD_TEMPLATE)
        
        # Replace placeholders in the card body
        for body_item in template.get("body", []):
            if "text" in body_item:
                body_item["text"] = body_item["text"].replace("${userName}", member_name)
        
        # Replace placeholders in the msteams entities
        if "msteams" in template and "entities" in template["msteams"]:
            for entity in template["msteams"]["entities"]:
                if "text" in entity:
                    entity["text"] = entity["text"].replace("${userName}", member_name)
                if "mentioned" in entity:
                    entity["mentioned"]["id"] = entity["mentioned"]["id"].replace("${userUPN}", member_upn)
                    entity["mentioned"]["id"] = entity["mentioned"]["id"].replace("${userAAD}", member_aad)
                    entity["mentioned"]["name"] = entity["mentioned"]["name"].replace("${userName}", member_name)
        
        # Send the adaptive card
        adaptive_card_attachment = {
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": template
        }
        
        activity = MessageActivityInput(
            text="",
            attachments=[adaptive_card_attachment]
        )
        await ctx.send(activity)
        
    except Exception as e:
        if "MemberNotFoundInConversation" in str(e):
            await ctx.send("Member not found.")
        else:
            raise


async def get_member_info(ctx: ActivityContext[MessageActivity]) -> None:
    """Get and display information about the current user."""
    try:
        conversation_id = ctx.activity.conversation.id
        user_id = ctx.activity.from_.id
        
        # Try to get member details from the API
        try:
            members = await ctx.api.conversations.members(conversation_id).get_all()
            member = None
            for m in members:
                if m.id == user_id:
                    member = m
                    break
            
            if member:
                await ctx.send(f"You are: {member.name}")
            else:
                await ctx.send(f"You are: {ctx.activity.from_.name}")
        except Exception:
            # Fallback to activity sender info
            await ctx.send(f"You are: {ctx.activity.from_.name}")
            
    except Exception as e:
        if "MemberNotFoundInConversation" in str(e):
            await ctx.send("Member not found.")
        else:
            raise


async def delete_card_activity(ctx: ActivityContext[MessageActivity]) -> None:
    """Delete the card that triggered this action."""
    try:
        if hasattr(ctx.activity, 'reply_to_id') and ctx.activity.reply_to_id:
            await ctx.api.conversations.activities(
                ctx.activity.conversation.id
            ).delete(ctx.activity.reply_to_id)
        else:
            await ctx.send("No card to delete.")
    except Exception as e:
        await ctx.send(f"Could not delete the card: {str(e)}")


async def send_immersive_reader_card(ctx: ActivityContext[MessageActivity]) -> None:
    """Send an Immersive Reader Adaptive Card."""
    adaptive_card_attachment = {
        "contentType": "application/vnd.microsoft.card.adaptive",
        "content": IMMERSIVE_READER_CARD_TEMPLATE
    }
    
    activity = MessageActivityInput(
        text="",
        attachments=[adaptive_card_attachment]
    )
    await ctx.send(activity)


def remove_recipient_mention(text: str, recipient_name: Optional[str] = None) -> str:
    """Remove bot mention from the text."""
    if not text:
        return ""
    
    # Remove <at>...</at> tags
    cleaned = re.sub(r'<at>.*?</at>', '', text)
    # Remove extra whitespace
    cleaned = ' '.join(cleaned.split())
    return cleaned.strip()


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]) -> None:
    """Handle incoming message activities."""
    await ctx.reply(TypingActivityInput())
    
    # Get the text and clean it (remove bot mention)
    raw_text = ctx.activity.text or ""
    recipient_name = ctx.activity.recipient.name if hasattr(ctx.activity, 'recipient') and ctx.activity.recipient else None
    text = remove_recipient_mention(raw_text, recipient_name).lower()
    
    # Handle different commands
    if "mention me" in text:
        await mention_adaptive_card_activity(ctx)
        return
    
    if "update" in text or "updatecardaction" in text:
        await send_update_card(ctx)
        return
    
    if "who" in text or "whoami" in text:
        await get_member_info(ctx)
        return
    
    if "immersivereader" in text:
        await send_immersive_reader_card(ctx)
        return
    
    if "delete" in text or "deletecard" in text:
        await delete_card_activity(ctx)
        return
    
    # Default: send welcome card
    await send_welcome_card(ctx)


@app.on_conversation_update
async def handle_members_added(ctx: ActivityContext) -> None:
    """Handle when new members are added to the conversation."""
    members_added = getattr(ctx.activity, 'members_added', [])
    if not members_added:
        return
        
    recipient_id = ctx.activity.recipient.id if hasattr(ctx.activity, 'recipient') else None
    conversation_type = getattr(ctx.activity.conversation, 'conversation_type', None)
    
    for member in members_added:
        if member.id != recipient_id and conversation_type != 'personal':
            given_name = getattr(member, 'given_name', '')
            surname = getattr(member, 'surname', '')
            name = f"{given_name} {surname}".strip() or getattr(member, 'name', 'User')
            await ctx.send(f"Welcome to the team {name}")


if __name__ == "__main__":
    asyncio.run(app.start())
