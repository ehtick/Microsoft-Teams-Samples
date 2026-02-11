# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from microsoft_teams.api import (
    Attachment,
    MessageActivityInput,
)
from microsoft_teams.apps import ActivityContext
from microsoft_teams.cards import (
    AdaptiveCard,
    TextBlock,
    TextInput,
    OpenUrlAction,
    ShowCardAction,
    SubmitAction,
    ToggleVisibilityAction,
)


async def send_adaptive_card_actions(ctx: ActivityContext) -> None:
    """Send Adaptive Card with various actions."""
    # Innermost nested card
    innermost_card = AdaptiveCard(
        body=[
            TextBlock(text="**Welcome To New Card**"),
            TextBlock(text="This is your new card inside another card")
        ]
    )

    # Middle nested card with ShowCard action
    middle_card = AdaptiveCard(
        body=[
            TextBlock(text="This card's action will show another card")
        ],
        actions=[
            ShowCardAction(title="Action.ShowCard", card=innermost_card)
        ]
    )

    # Submit form card
    submit_form_card = AdaptiveCard(
        body=[
            TextInput(
                id="name",
                label="Please enter your name:",
                is_required=True,
                error_message="Name is required"
            )
        ],
        actions=[
            SubmitAction(title="Submit")
        ]
    )

    # Main card with all actions
    card = AdaptiveCard(
        body=[
            TextBlock(text="Adaptive Card Actions")
        ],
        actions=[
            OpenUrlAction(title="Action Open URL", url="https://adaptivecards.io"),
            ShowCardAction(title="Action Submit", card=submit_form_card),
            ShowCardAction(title="Action ShowCard", card=middle_card)
        ]
    )
    
    card_payload = card.model_dump(by_alias=True, exclude_none=True)
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card_payload
    )
    await ctx.send(MessageActivityInput(attachments=[attachment]))


async def send_toggle_visibility_card(ctx: ActivityContext) -> None:
    """Send Toggle Visibility Card."""
    card = AdaptiveCard(
        body=[
            TextBlock(
                text="**Action.ToggleVisibility example**: click the button to show or hide a welcome message"
            ),
            TextBlock(
                id="helloWorld",
                is_visible=False,
                text="**Hello World!**",
                size="ExtraLarge" 
            )
        ],
        actions=[
            ToggleVisibilityAction(
                title="Click me!",
                target_elements=["helloWorld"]
            )
        ]
    )
    
    card_payload = card.model_dump(by_alias=True, exclude_none=True)
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card_payload
    )
    await ctx.send(MessageActivityInput(attachments=[attachment]))
