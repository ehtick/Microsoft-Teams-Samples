"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

import asyncio
import re
from typing import Optional

from dotenv import load_dotenv
from microsoft_teams.api import MessageActivity, TypingActivityInput, MessageActivityInput, InstalledActivity
from microsoft_teams.apps import ActivityContext, App

# Load environment variables
load_dotenv()

app = App()

# Simple in-memory storage for conversation references (for proactive messaging)
# In production, use persistent storage like a database
conversation_storage: dict[str, str] = {}


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
        if "MemberNotFoundInConversation" not in str(e):
            raise


async def mention_user(ctx: ActivityContext[MessageActivity]) -> None:
    """Mention the current user in a text message."""
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
            
            if not member:
                member = ctx.activity.from_
        except Exception:
            member = ctx.activity.from_
        
        member_name = member.name if hasattr(member, 'name') and member.name else "User"
        member_id = member.id
        
        # Create the mention text
        mention_text = f"<at>{member_name}</at>"
        
        # Create the mention entity
        mention_entity = {
            "type": "mention",
            "text": mention_text,
            "mentioned": {
                "id": member_id,
                "name": member_name
            }
        }
        
        # Send message with mention entity
        activity = MessageActivityInput(
            text=f"Hello {mention_text}!",
            entities=[mention_entity]
        )
        await ctx.send(activity)
        
    except Exception as e:
        if "MemberNotFoundInConversation" not in str(e):
            raise


def remove_recipient_mention(text: str, recipient_name: Optional[str] = None) -> str:
    """Remove bot mention from the text."""
    if not text:
        return ""
    
    # Remove <at>...</at> tags
    cleaned = re.sub(r'<at>.*?</at>', '', text)
    # Remove extra whitespace
    cleaned = ' '.join(cleaned.split())
    return cleaned.strip()


async def send_proactive_notification(user_id: str, message: str = "Hey! This is a proactive message from the bot!") -> bool:
    """Send a proactive message to a user."""
    
    conversation_id = conversation_storage.get(user_id, "")
    if not conversation_id:
        return False
    
    activity = MessageActivityInput(text=message)
    await app.send(conversation_id, activity)
    return True


async def delayed_proactive_message(user_id: str, delay_seconds: int = 10) -> None:
    """Send a proactive message after a delay."""

    await asyncio.sleep(delay_seconds)
    await send_proactive_notification(
        user_id, 
        f"Reminder: This proactive message was sent {delay_seconds} seconds after your request!"
    )

@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]) -> None:
    """Handle incoming message activities."""
    await ctx.reply(TypingActivityInput())
    
    # Get the text and clean it (remove bot mention)
    raw_text = ctx.activity.text or ""
    recipient_name = ctx.activity.recipient.name if hasattr(ctx.activity, 'recipient') and ctx.activity.recipient else None
    text = remove_recipient_mention(raw_text, recipient_name).lower()
    
    # Store conversation reference for proactive messaging (from any message)
    user_aad_id = ctx.activity.from_.aad_object_id
    if user_aad_id:
        conversation_storage[user_aad_id] = ctx.activity.conversation.id
    
    # Handle proactive messaging command
    if "remind" in text or "proactive" in text:
        if user_aad_id:
            await ctx.send("Got it! I'll send you a proactive message in 10 seconds...")
            # Schedule the proactive message (runs in background)
            asyncio.create_task(delayed_proactive_message(user_aad_id, 10))
        else:
            await ctx.send("Sorry, I couldn't identify your user ID for proactive messaging.")
        return
    
    # Handle different commands
    if "mentionme" in text:
        await mention_user(ctx)
        return
    
    if "who" in text or "whoami" in text:
        await get_member_info(ctx)
        return
    
    # Echo back any other message
    if text:
        await ctx.send(f"Echo: {text}")


@app.on_conversation_update
async def handle_members_added(ctx: ActivityContext) -> None:
    """Handle when new members are added to the conversation."""
    members_added = getattr(ctx.activity, 'members_added', [])
    if not members_added:
        return
        
    recipient_id = ctx.activity.recipient.id if hasattr(ctx.activity, 'recipient') else None
    conversation_type = getattr(ctx.activity.conversation, 'conversation_type', None)
    
    for member in members_added:
        # If the bot itself was added, send welcome message
        if member.id == recipient_id:
            welcome_message = (
                "Welcome to Microsoft Teams conversationUpdate events demo bot.\n\n"
                "Available commands:\n"
                "- **mention me** - Bot will mention you in the reply\n"
                "- **whoami** - Get your user information\n"
                "- **remind me** or **proactive** - Bot will send a proactive message in 10 seconds\n"
                "- **Echo** - Bot will echo your message back"
            )
            await ctx.send(welcome_message)
        # If another member was added to a non-personal conversation
        elif conversation_type != 'personal':
            given_name = getattr(member, 'given_name', '')
            surname = getattr(member, 'surname', '')
            name = f"{given_name} {surname}".strip() or getattr(member, 'name', 'User')
            await ctx.send(f"Welcome to the team {name}")


if __name__ == "__main__":
    asyncio.run(app.start())
