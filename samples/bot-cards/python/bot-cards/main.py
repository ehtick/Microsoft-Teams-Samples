# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import asyncio

from dotenv import load_dotenv
from microsoft_teams.api import (
    InstalledActivity,
    InvokeActivity,
    InvokeResponse,
    MessageActivity,
)
from microsoft_teams.apps import ActivityContext, App

from handlers.adaptive_card import send_adaptive_card_actions, send_toggle_visibility_card
from handlers.attachments import (
    get_value,
    pending_uploads,
    send_file_card,
    file_upload_handler,
    handle_file_download,
    process_inline_image,
)

load_dotenv()

app = App()


async def send_welcome_message(ctx: ActivityContext) -> None:
    """Sends welcome message with available commands."""
    await ctx.send("Welcome to the Teams Bot Cards!")


@app.on_install_add
async def handle_install(ctx: ActivityContext[InstalledActivity]) -> None:
    """Handle membersAdded event to send welcome message."""
    await send_welcome_message(ctx)


@app.on_invoke
async def handle_invoke(ctx: ActivityContext[InvokeActivity]) -> InvokeResponse:
    """Handle file consent invoke activities."""
    activity = ctx.activity

    if activity.name != "fileConsent/invoke":
        return InvokeResponse(status=200)

    response = activity.value
    action = get_value(response, "action")
    context = get_value(response, "context")
    filename = get_value(context, "filename")
    file_id = get_value(context, "file_id")

    if action == "accept":
        # Start upload in background task to avoid timeout
        asyncio.create_task(file_upload_handler(ctx, response, filename, file_id))

    elif action == "decline":
        # Clean up cached content if any
        if file_id and file_id in pending_uploads:
            del pending_uploads[file_id]
        await ctx.send(f"Declined. We won't upload file <b>{filename}</b>.")

    return InvokeResponse(status=200)


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]) -> None:
    """Handle incoming messages."""
    attachments = ctx.activity.attachments or []

    # Handle file attachments
    if attachments:
        attachment = attachments[0]
        content_type = attachment.content_type or ""

        if content_type == "application/vnd.microsoft.teams.file.download.info":
            await handle_file_download(ctx, attachment)
            return

        # Handle any inline attachment (images, gifs, videos, etc.)
        if attachment.content_url:
            await process_inline_image(ctx, attachment)
            return

    text = (ctx.activity.text or "").lower()

    # Handle Adaptive Card Actions commands
    if "card actions" in text:
        await send_adaptive_card_actions(ctx)
        return

    if "togglevisibility" in text:
        await send_toggle_visibility_card(ctx)
        return

    # Handle card submit data
    if ctx.activity.value:
        name = ctx.activity.value.get("name", "N/A") if isinstance(ctx.activity.value, dict) else "N/A"
        await ctx.send(f"Data Submitted: {name}")
        return

    # Default: trigger file upload for "send file" or similar
    if "send file" in text or "file" in text:
        await send_file_card(ctx, "teams-logo.png")
        return

    # Welcome message for unrecognized commands
    await send_welcome_message(ctx)


if __name__ == "__main__":
    asyncio.run(app.start())