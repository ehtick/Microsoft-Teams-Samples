// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using AdaptiveCards;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api;
using System.Text.Json;

namespace Microsoft.Teams.Samples.BotCards.Handlers;

public static class Cards
{
    // Send Adaptive Card with various actions
    public static async Task SendAdaptiveCardActions(dynamic context)
    {
        var adaptiveCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 4))
        {
            Body = new List<AdaptiveElement>
            {
                new AdaptiveTextBlock
                {
                    Text = "Adaptive Card Actions",
                    Weight = AdaptiveTextWeight.Bolder,
                    Size = AdaptiveTextSize.Large
                }
            },
            Actions = new List<AdaptiveAction>
            {
                new AdaptiveOpenUrlAction
                {
                    Title = "Action Open URL",
                    Url = new Uri("https://adaptivecards.io")
                },
                new AdaptiveShowCardAction
                {
                    Title = "Action Submit",
                    Card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 4))
                    {
                        Body = new List<AdaptiveElement>
                        {
                            new AdaptiveTextInput
                            {
                                Id = "name",
                                Label = "Please enter your name:",
                                IsRequired = true,
                                ErrorMessage = "Name is required"
                            }
                        },
                        Actions = new List<AdaptiveAction>
                        {
                            new AdaptiveSubmitAction
                            {
                                Title = "Submit"
                            }
                        }
                    }
                },
                new AdaptiveShowCardAction
                {
                    Title = "Action ShowCard",
                    Card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 4))
                    {
                        Body = new List<AdaptiveElement>
                        {
                            new AdaptiveTextBlock
                            {
                                Text = "This card's action will show another card"
                            }
                        },
                        Actions = new List<AdaptiveAction>
                        {
                            new AdaptiveShowCardAction
                            {
                                Title = "Action.ShowCard",
                                Card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 4))
                                {
                                    Body = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            Text = "Welcome To New Card"
                                        }
                                    },
                                    Actions = new List<AdaptiveAction>
                                    {
                                        new AdaptiveSubmitAction
                                        {
                                            Title = "Click Me",
                                            Data = new { value = "Button has Clicked" }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        };

        var message = new MessageActivity();
        message.Attachments = new List<Attachment>
        {
            new Attachment
            {
                ContentType = new ContentType("application/vnd.microsoft.card.adaptive"),
                Content = JsonSerializer.Deserialize<JsonElement>(adaptiveCard.ToJson())
            }
        };
        await context.Send(message);
    }

    // Send Toggle Visibility Card
    public static async Task SendToggleVisibilityCard(dynamic context)
    {
        var adaptiveCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 4))
        {
            Body = new List<AdaptiveElement>
            {
                new AdaptiveTextBlock
                {
                    Text = "**Action.ToggleVisibility example**: click the button to show or hide a welcome message",
                    Wrap = true
                },
                new AdaptiveTextBlock
                {
                    Text = "**Hello World!**",
                    Id = "helloWorld",
                    IsVisible = false,
                    Size = AdaptiveTextSize.ExtraLarge
                }
            },
            Actions = new List<AdaptiveAction>
            {
                new AdaptiveToggleVisibilityAction
                {
                    Title = "Click me!",
                    TargetElements = new List<AdaptiveTargetElement>
                    {
                        new AdaptiveTargetElement("helloWorld")
                    }
                }
            }
        };

        var message = new MessageActivity();
        message.Attachments = new List<Attachment>
        {
            new Attachment
            {
                ContentType = new ContentType("application/vnd.microsoft.card.adaptive"),
                Content = JsonSerializer.Deserialize<JsonElement>(adaptiveCard.ToJson())
            }
        };
        await context.Send(message);
    }
}
