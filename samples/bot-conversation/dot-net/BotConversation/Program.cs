// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Teams.Plugins.AspNetCore.Extensions;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api.Clients;
using System.Collections.Concurrent;
using System.Text.Json;

// Static state for tracking messages
var counter = 0;
var userLastMessageIds = new ConcurrentDictionary<string, string>();
var builder = WebApplication.CreateBuilder(args);
builder.AddTeams();
var webApp = builder.Build();
var configuration = webApp.Services.GetRequiredService<IConfiguration>();
var teamsApp = webApp.UseTeams(true);

// Main message handler
teamsApp.OnMessage(async context =>
{
    var text = context.Activity.Text?.Trim().ToLower() ?? "";
    if (text.Contains("mentionme") || text.Contains("mention me"))
        await MentionAdaptiveCardActivityAsync(context);
    else if (text.Contains("whoami"))
        await GetSingleMemberAsync(context);
    else if (text.Contains("update"))
        await UpdateCardActivityAsync(context);
    else if (text.Contains("delete"))
        await DeleteCardActivityAsync(context);
    else if (text.Contains("immersivereader"))
        await SendImmersiveReaderCardAsync(context);
    else
        await SendWelcomeCard(context);
});

// Conversation members added handler
teamsApp.OnConversationUpdate(async context =>
{
    if (context.Activity.MembersAdded != null)
    {
        foreach (var member in context.Activity.MembersAdded)
        {
            if (member.Id != context.Activity.Recipient.Id)
            {
                await context.Send($"Welcome to the team {member.Name}! This bot demonstrates Teams conversation events and adaptive cards. Try 'show welcome' to get started.");
            }
        }
    }
});

webApp.Run();

// Helper methods
async Task MentionAdaptiveCardActivityAsync(IContext<MessageActivity> context)
{
    try
    {
        var userName = context.Activity.From.Name ?? "User";
        var userId = context.Activity.From.AadObjectId ?? context.Activity.From.Id;
        
        var cardContent = new
        {
            type = "AdaptiveCard",
            version = "1.5",
            body = new[]
            {
                new
                {
                    type = "TextBlock",
                    text = $"Mention a user by User Principle Name: Hello <at>{userName} UPN</at>"
                },
                new
                {
                    type = "TextBlock",
                    text = $"Mention a user by AAD Object Id: Hello <at>{userName} AAD</at>"
                }
            },
            msteams = new
            {
                entities = new[]
                {
                    new
                    {
                        type = "mention",
                        text = $"<at>{userName} UPN</at>",
                        mentioned = new
                        {
                            id = userId,
                            name = userName
                        }
                    },
                    new
                    {
                        type = "mention",
                        text = $"<at>{userName} AAD</at>",
                        mentioned = new
                        {
                            id = userId,
                            name = userName
                        }
                    }
                }
            }
        };

        var cardElement = JsonSerializer.SerializeToElement(cardContent);

        var messageActivity = new MessageActivity
        {
            Attachments = new List<Microsoft.Teams.Api.Attachment>
            {
                new Microsoft.Teams.Api.Attachment
                {
                    ContentType = Microsoft.Teams.Api.ContentType.AdaptiveCard,
                    Content = cardElement
                }
            }
        };
        
        await context.Send(messageActivity);
    }
    catch (Exception ex)
    {
        await context.Send($"Error sending mention card: {ex.Message}");
    }
}

async Task GetSingleMemberAsync(IContext<MessageActivity> context)
{
    try
    {
        var members = await context.Api.Conversations.Members.GetAsync(context.Activity.Conversation.Id);
        var member = members?.FirstOrDefault(m => m.Id == context.Activity.From.Id);
        
        if (member != null)
        {
            await context.Send($"You are: {member.Name}.");
        }
        else
        {
            await context.Send("Unable to find your member information.");
        }
    }
    catch (Exception ex)
    {
        await context.Send($"Unable to get member information: {ex.Message}");
    }
}

async Task UpdateCardActivityAsync(IContext<MessageActivity> context)
{
    try
    {
        var userId = context.Activity.From.Id;
        
        if (!userLastMessageIds.TryGetValue(userId, out var messageId) || string.IsNullOrEmpty(messageId))
        {
            await context.Send("No card found to update. Please send a message first, then try 'update'.");
            return;
        }

        counter++;
        var heroCardContent = new
        {
            title = "Updated card",
            text = $"Update count {counter}",
            buttons = new[]
            {
                new { type = "imBack", title = "Who am I?", value = "whoami" },
                new { type = "imBack", title = "Find me in Adaptive Card", value = "mention me" },
                new { type = "imBack", title = "Delete card", value = "delete" },
                new { type = "imBack", title = "Send Immersive Reader Card", value = "immersivereader" },
                new { type = "imBack", title = "Update Card", value = "update" }
            }
        };

        var messageActivity = new MessageActivity
        {
            Type = ActivityType.Message,
            Attachments = new List<Microsoft.Teams.Api.Attachment>
            {
                new Microsoft.Teams.Api.Attachment
                {
                    ContentType = Microsoft.Teams.Api.ContentType.HeroCard,
                    Content = JsonSerializer.SerializeToElement(heroCardContent)
                }
            }
        };
        
        await context.Api.Conversations.Activities.UpdateAsync(
            context.Activity.Conversation.Id,
            messageId,
            messageActivity
        );
    }
    catch (Exception ex)
    {
        await context.Send($"Unable to update card: {ex.Message}");
    }
}

async Task DeleteCardActivityAsync(IContext<MessageActivity> context)
{
    try
    {
        var userId = context.Activity.From.Id;
        
        if (!userLastMessageIds.TryGetValue(userId, out var messageId) || string.IsNullOrEmpty(messageId))
        {
            await context.Send("No card found to delete. Please send 'welcome' first, then try 'delete'.");
            return;
        }

        await context.Api.Conversations.Activities.DeleteAsync(
            context.Activity.Conversation.Id,
            messageId
        );
        userLastMessageIds.TryRemove(userId, out _);
        await context.Send("Card has been deleted successfully.");
    }
    catch (Exception ex)
    {
        await context.Send($"Unable to delete card: {ex.Message}");
    }
}

async Task SendWelcomeCard(IContext<MessageActivity> context)
{
    counter++;
    var heroCardContent = new
    {
        title = "Updated card",
        text = $"Update count {counter}",
        buttons = new[]
        {
            new { type = "imBack", title = "Who am I?", value = "whoami" },
            new { type = "imBack", title = "Find me in Adaptive Card", value = "mention me" },
            new { type = "imBack", title = "Delete card", value = "delete" },
            new { type = "imBack", title = "Send Immersive Reader Card", value = "immersivereader" },
            new { type = "imBack", title = "Update Card", value = "update" }
        }
    };
    var messageActivity = new MessageActivity
    {
        Type = ActivityType.Message,
        Attachments = new List<Microsoft.Teams.Api.Attachment>
        {
            new Microsoft.Teams.Api.Attachment
            {
                ContentType = Microsoft.Teams.Api.ContentType.HeroCard,
                Content = JsonSerializer.SerializeToElement(heroCardContent)
            }
        }
    };    
    var response = await context.Send(messageActivity);
    if (response != null && !string.IsNullOrEmpty(response.Id))
    {
        userLastMessageIds[context.Activity.From.Id] = response.Id;
    }
}

async Task SendImmersiveReaderCardAsync(IContext<MessageActivity> context)
{
    try
    {
        var cardContent = new
        {
            type = "AdaptiveCard",
            version = "1.5",
            body = new object[]
            {
                new
                {
                    type = "ColumnSet",
                    columns = new object[]
                    {
                        new
                        {
                            type = "Column",
                            width = "auto",
                            items = new object[]
                            {
                                new
                                {
                                    type = "Image",
                                    size = "Small",
                                    url = "https://adaptivecards.io/content/airplane.png",
                                    altText = "Airplane"
                                }
                            }
                        },
                        new
                        {
                            type = "Column",
                            width = "stretch",
                            items = new object[]
                            {
                                new
                                {
                                    type = "TextBlock",
                                    text = "Flight Status",
                                    horizontalAlignment = "Right",
                                    isSubtle = true,
                                    wrap = true
                                },
                                new
                                {
                                    type = "TextBlock",
                                    text = "DELAYED",
                                    horizontalAlignment = "Right",
                                    spacing = "None",
                                    size = "Large",
                                    color = "Attention",
                                    wrap = true
                                }
                            }
                        }
                    }
                },
                new
                {
                    type = "ColumnSet",
                    separator = true,
                    spacing = "Medium",
                    columns = new object[]
                    {
                        new
                        {
                            type = "Column",
                            width = "stretch",
                            items = new object[]
                            {
                                new
                                {
                                    type = "TextBlock",
                                    text = "Passengers",
                                    isSubtle = true,
                                    weight = "Bolder",
                                    wrap = true
                                },
                                new { type = "TextBlock", text = "Sarah Hum", spacing = "Small", wrap = true },
                                new { type = "TextBlock", text = "Jeremy Goldberg", spacing = "Small", wrap = true },
                                new { type = "TextBlock", text = "Evan Litvak", spacing = "Small", wrap = true }
                            }
                        },
                        new
                        {
                            type = "Column",
                            width = "auto",
                            items = new object[]
                            {
                                new
                                {
                                    type = "TextBlock",
                                    text = "Seat",
                                    horizontalAlignment = "Right",
                                    isSubtle = true,
                                    weight = "Bolder",
                                    wrap = true
                                },
                                new { type = "TextBlock", text = "14A", horizontalAlignment = "Right", spacing = "Small", wrap = true },
                                new { type = "TextBlock", text = "14B", horizontalAlignment = "Right", spacing = "Small", wrap = true },
                                new { type = "TextBlock", text = "14C", horizontalAlignment = "Right", spacing = "Small", wrap = true }
                            }
                        }
                    }
                },
                new
                {
                    type = "ColumnSet",
                    spacing = "Medium",
                    separator = true,
                    columns = new object[]
                    {
                        new
                        {
                            type = "Column",
                            width = 1,
                            items = new object[]
                            {
                                new { type = "TextBlock", text = "Flight", isSubtle = true, weight = "Bolder", wrap = true },
                                new { type = "TextBlock", text = "KL0605", spacing = "Small", wrap = true }
                            }
                        },
                        new
                        {
                            type = "Column",
                            width = 1,
                            items = new object[]
                            {
                                new { type = "TextBlock", text = "Departs", isSubtle = true, horizontalAlignment = "Center", weight = "Bolder", wrap = true },
                                new { type = "TextBlock", text = "10:10 AM", color = "Attention", weight = "Bolder", horizontalAlignment = "Center", spacing = "Small", wrap = true }
                            }
                        },
                        new
                        {
                            type = "Column",
                            width = 1,
                            items = new object[]
                            {
                                new { type = "TextBlock", text = "Arrives", isSubtle = true, horizontalAlignment = "Right", weight = "Bolder", wrap = true },
                                new { type = "TextBlock", text = "12:00 AM", color = "Attention", horizontalAlignment = "Right", weight = "Bolder", spacing = "Small", wrap = true }
                            }
                        }
                    }
                },
                new
                {
                    type = "ColumnSet",
                    spacing = "Medium",
                    separator = true,
                    columns = new object[]
                    {
                        new
                        {
                            type = "Column",
                            width = 1,
                            items = new object[]
                            {
                                new { type = "TextBlock", text = "Amsterdam Airport", isSubtle = true, wrap = true },
                                new { type = "TextBlock", text = "AMS", size = "ExtraLarge", color = "Accent", spacing = "None", wrap = true }
                            }
                        },
                        new
                        {
                            type = "Column",
                            width = "auto",
                            items = new object[]
                            {
                                new { type = "TextBlock", text = " ", wrap = true },
                                new { type = "Image", url = "https://adaptivecards.io/content/airplane.png", altText = "Airplane", size = "Small" }
                            }
                        },
                        new
                        {
                            type = "Column",
                            width = 1,
                            items = new object[]
                            {
                                new { type = "TextBlock", text = "San Francisco Airport", isSubtle = true, horizontalAlignment = "Right", wrap = true },
                                new { type = "TextBlock", text = "SFO", horizontalAlignment = "Right", size = "ExtraLarge", color = "Accent", spacing = "None", wrap = true }
                            }
                        }
                    }
                }
            }
        };        
        var cardElement = JsonSerializer.SerializeToElement(cardContent);        
        var messageActivity = new MessageActivity
        {
            Attachments = new List<Microsoft.Teams.Api.Attachment>
            {
                new Microsoft.Teams.Api.Attachment
                {
                    ContentType = Microsoft.Teams.Api.ContentType.AdaptiveCard,
                    Content = cardElement
                }
            }
        };        
        await context.Send(messageActivity);
    }
    catch (Exception ex)
    {
        await context.Send($"Error sending immersive reader card: {ex.Message}");
    }
}