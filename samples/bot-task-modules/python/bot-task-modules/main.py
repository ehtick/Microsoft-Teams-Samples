# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import asyncio
import os

from dotenv import load_dotenv
from microsoft_teams.apps import App, ActivityContext
from microsoft_teams.api import (
    MessageActivity,
    MessageActivityInput,
    TaskFetchInvokeActivity,
    TaskSubmitInvokeActivity,
    InvokeResponse,
    card_attachment,
    AdaptiveCardAttachment,
    TaskModuleResponse,
    TaskModuleContinueResponse,
    TaskModuleMessageResponse,
    UrlTaskModuleTaskInfo,
    CardTaskModuleTaskInfo,
    HeroCard,
    HeroCardAttachment,
    CardAction,
)
from microsoft_teams.cards import AdaptiveCard, TextBlock, SubmitAction, TextInput, TaskFetchAction
from task_module_config import TaskModuleIds


# Load environment variables
load_dotenv()

# Get BASE_URL from environment
BASE_URL = os.environ.get("BaseUrl", "")

# Initialize Teams App
app = App()

# Host static webpages for task modules
app.page("youtube", os.path.join(os.path.dirname(__file__), "pages", "YouTube"), "/youtube")
app.page("customform", os.path.join(os.path.dirname(__file__), "pages", "CustomForm"), "/customform")

# Mount CSS files for task module pages
app.http.mount("css", os.path.join(os.path.dirname(__file__), "pages", "css"))


def create_hero_card_attachment() -> HeroCardAttachment:
    """Creates a HeroCard with task module options."""
    hero_card = HeroCard(
        title="Task Module Invocation from Hero Card",
        buttons=[
            CardAction.model_construct(
                type="invoke",
                title=TaskModuleIds.ADAPTIVE_CARD.button_title,
                value={"type": "task/fetch", "data": TaskModuleIds.ADAPTIVE_CARD.id}
            ),
            CardAction.model_construct(
                type="invoke",
                title=TaskModuleIds.CUSTOM_FORM.button_title,
                value={"type": "task/fetch", "data": TaskModuleIds.CUSTOM_FORM.id}
            ),
            CardAction.model_construct(
                type="invoke",
                title=TaskModuleIds.YOUTUBE.button_title,
                value={"type": "task/fetch", "data": TaskModuleIds.YOUTUBE.id}
            )
        ]
    )
    return HeroCardAttachment(content=hero_card)


def create_adaptive_card_with_task_module_options() -> AdaptiveCard:
    """Creates an AdaptiveCard with task module options."""
    card = AdaptiveCard(version="1.4").with_body([
        TextBlock(text="Task Module Invocation from Adaptive Card", weight="Bolder", size="Large")
    ]).with_actions([
        TaskFetchAction(value={"data": TaskModuleIds.ADAPTIVE_CARD.id}).with_title(TaskModuleIds.ADAPTIVE_CARD.button_title),
        TaskFetchAction(value={"data": TaskModuleIds.CUSTOM_FORM.id}).with_title(TaskModuleIds.CUSTOM_FORM.button_title),
        TaskFetchAction(value={"data": TaskModuleIds.YOUTUBE.id}).with_title(TaskModuleIds.YOUTUBE.button_title)
    ])
    
    return card


def create_adaptive_card_for_task_module() -> AdaptiveCard:
    """Creates an AdaptiveCard to be shown in a task module."""
    return AdaptiveCard(
        schema="http://adaptivecards.io/schemas/adaptive-card.json",
        version="1.0"
    ).with_body([
        TextBlock(text="Enter Text Here", weight="Bolder"),
        TextInput()
            .with_id("usertext")
            .with_placeholder("add some text and submit")
            .with_is_multiline(True)
    ]).with_actions([
        SubmitAction().with_title("Submit")
    ])


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]) -> None:
    """Handles incoming messages and displays two cards: a HeroCard and an AdaptiveCard."""
    # Create hero card attachment
    hero_card = create_hero_card_attachment()
    
    # Create adaptive card
    adaptive_card = create_adaptive_card_with_task_module_options()
    
    # Send both cards - use add_attachments for HeroCard and add_card for AdaptiveCard
    message = MessageActivityInput().add_attachments(hero_card).add_card(adaptive_card)
    await ctx.send(message)


@app.on_dialog_open
async def handle_task_module_fetch(ctx: ActivityContext[TaskFetchInvokeActivity]) -> InvokeResponse:
    """Called when the user selects an option from the displayed HeroCard or AdaptiveCard."""
    # Get the task module data from the request
    value = ctx.activity.value
    card_data = None
    
    # Extract data - try multiple approaches
    if hasattr(value, "data") and value.data:
        data = value.data
        if isinstance(data, dict):
            card_data = data.get("data")
        elif isinstance(data, str):
            card_data = data
    
    # If card_data is still None, try to get it from the raw value
    if card_data is None and hasattr(value, "data") and isinstance(value.data, dict):
        # Maybe the data key contains the ID directly without nesting
        if "type" in value.data and value.data.get("type") == "task/fetch":
            card_data = value.data.get("data")
    
    # Default to AdaptiveCard if we can't determine what was clicked
    if card_data is None:
        card_data = TaskModuleIds.ADAPTIVE_CARD.id
    
    task_info = None
    
    if card_data == TaskModuleIds.YOUTUBE.id:
        task_info = UrlTaskModuleTaskInfo(
            title=TaskModuleIds.YOUTUBE.title,
            width=TaskModuleIds.YOUTUBE.width,
            height=TaskModuleIds.YOUTUBE.height,
            url=f"{BASE_URL}/youtube",
            fallback_url=f"{BASE_URL}/youtube"
        )
    elif card_data == TaskModuleIds.CUSTOM_FORM.id:
        task_info = UrlTaskModuleTaskInfo(
            title=TaskModuleIds.CUSTOM_FORM.title,
            width=TaskModuleIds.CUSTOM_FORM.width,
            height=TaskModuleIds.CUSTOM_FORM.height,
            url=f"{BASE_URL}/customform",
            fallback_url=f"{BASE_URL}/customform"
        )
    else:  # Default to ADAPTIVE_CARD
        task_info = CardTaskModuleTaskInfo(
            title=TaskModuleIds.ADAPTIVE_CARD.title,
            width=TaskModuleIds.ADAPTIVE_CARD.width,
            height=TaskModuleIds.ADAPTIVE_CARD.height,
            card=card_attachment(AdaptiveCardAttachment(content=create_adaptive_card_for_task_module()))
        )
    
    # Use body=TaskModuleResponse as per the documentation
    return InvokeResponse(
        body=TaskModuleResponse(
            task=TaskModuleContinueResponse(value=task_info)
        )
    )


@app.on_dialog_submit
async def handle_task_module_submit(ctx: ActivityContext[TaskSubmitInvokeActivity]) -> InvokeResponse:
    """Called when data is being returned from the selected option."""
    # Get the submitted data
    value = ctx.activity.value
    data = value.data if hasattr(value, "data") else {}
    
    # Build a nicely formatted Adaptive Card to display the submitted data
    body_items = [
        TextBlock(text="Task Module Submission Received", size="Large", weight="Bolder")
    ]
    
    # Add each field from the submitted data
    if data:
        for key, val in data.items():
            formatted_key = key.replace("_", " ").title()
            body_items.append(
                TextBlock(text=f"**{formatted_key}:** {val}", wrap=True)
            )
    else:
        body_items.append(TextBlock(text="No data submitted", isSubtle=True))
    
    result_card = AdaptiveCard(
        schema="http://adaptivecards.io/schemas/adaptive-card.json"
    ).with_body(body_items)
    
    # Send the formatted card
    message = MessageActivityInput(text="Task module submission received")
    message.add_card(result_card)
    await ctx.send(message)
    
    # Return a message response
    return InvokeResponse(
        body=TaskModuleResponse(
            task=TaskModuleMessageResponse(value="Thanks!")
        )
    )


# Starts the Teams bot application and listens for incoming requests
if __name__ == "__main__":
    asyncio.run(app.start())