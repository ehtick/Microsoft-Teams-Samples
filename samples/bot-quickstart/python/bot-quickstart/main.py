# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import asyncio

from dotenv import load_dotenv
from microsoft_teams.api import MessageActivity, MessageActivityInput
from microsoft_teams.apps import ActivityContext, App

# Load environment variables
load_dotenv()

# Initialize Teams App - automatically uses CLIENT_ID and CLIENT_SECRET from environment variables
# Note: .env file is only required when running on Teams (not needed for local development with devtools)
app = App()

# Simple in-memory storage for conversation references (for proactive messaging)
# In production, use persistent storage like a database
conversation_storage: dict[str, str] = {}


async def send_proactive_notification(
    user_id: str, 
    message: str = "Hey! This is a proactive message from the bot!"
) -> bool:
    """Send a proactive message to a user."""
    conversation_id = conversation_storage.get(user_id)
    if not conversation_id:
        return False
    
    await app.send(conversation_id, MessageActivityInput(text=message))
    return True


async def delayed_proactive_message(user_id: str, delay_seconds: int = 10) -> None:
    """Send a proactive message after a delay."""
    await asyncio.sleep(delay_seconds)
    await send_proactive_notification(
        user_id,
        f"Reminder: This proactive message was sent {delay_seconds} seconds after your request!"
    )


async def send_welcome_message(ctx: ActivityContext) -> None:
    """Sends a welcome message with available commands."""
    welcome_message = (
        "Welcome to the Teams Quickstart Bot!\n\n"
        "Available commands:\n"
        "- **mention me** - Bot will mention you in the reply\n"
        "- **whoami** - Get your user information\n"
        "- **proactive** - Bot will send a proactive message in 10 seconds\n"
        "- **echo** - Bot will echo back your message"
    )
    await ctx.send(MessageActivityInput(text=welcome_message))


async def echo_message(ctx: ActivityContext, text: str) -> None:
    """Echo back the user's message."""
    await ctx.send(MessageActivityInput(text=f"**Echo:** {text}"))


async def get_single_member(ctx: ActivityContext[MessageActivity]) -> None:
    """Retrieves and displays information about the current user."""
    try:
        # Get user info directly from the activity
        user = ctx.activity.from_
        await ctx.send(MessageActivityInput(text=f"You are: {user.name}"))
    except Exception as error:
        print(f"Error getting member: {error}")


async def mention_user(ctx: ActivityContext[MessageActivity]) -> None:
    """Mention a user in a message."""
    try:
        # Get user info directly from the activity
        user = ctx.activity.from_
        user_id = user.id
        user_name = user.name
        
        # Create a text message with user mention
        mention_text = f"<at>{user_name}</at>"
        await ctx.send(MessageActivityInput(
            text=f"Hello {mention_text}",
            entities=[
                {
                    "type": "mention",
                    "text": mention_text,
                    "mentioned": {
                        "id": user_id,
                        "name": user_name,
                        "role": "user"
                    }
                }
            ]
        ))
    except Exception as error:
        print(f"Error mentioning user: {error}")


@app.on_conversation_update
async def handle_conversation_update(ctx: ActivityContext) -> None:
    """Handle conversation update events (when bot is added or members join)."""
    members_added = getattr(ctx.activity, 'members_added', [])
    
    for member in members_added:
        # Check if bot was added to the conversation
        if member.id == ctx.activity.recipient.id:
            await send_welcome_message(ctx)


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]) -> None:
    """Handles incoming messages and routes to appropriate functions based on message content."""
    # Get message text and normalize it
    text = (ctx.activity.text or "").strip().lower()
    
    # Store conversation reference for proactive messaging (from any message)
    user_aad_id = ctx.activity.from_.aad_object_id
    if user_aad_id:
        conversation_storage[user_aad_id] = ctx.activity.conversation.id
    
    # Handle proactive messaging command
    if "proactive" in text:
        if user_aad_id:
            await ctx.send(MessageActivityInput(
                text="Got it! I'll send you a proactive message in 10 seconds..."
            ))
            # Schedule the proactive message (runs in background)
            asyncio.create_task(delayed_proactive_message(user_aad_id, 10))
        else:
            await ctx.send(MessageActivityInput(
                text="Sorry, I couldn't identify your user ID for proactive messaging."
            ))
        return
    
    # Handle mention me command
    if "mentionme" in text or "mention me" in text:
        await mention_user(ctx)
    # Handle whoami command
    elif "whoami" in text:
        await get_single_member(ctx)
    # Handle hi/hello - echo back
    elif "hi" in text or "hello" in text:
        await echo_message(ctx, text)
    # Default: send welcome message
    else:
        await send_welcome_message(ctx)


# Starts the Teams bot application and listens for incoming requests
if __name__ == "__main__":
    asyncio.run(app.start())
