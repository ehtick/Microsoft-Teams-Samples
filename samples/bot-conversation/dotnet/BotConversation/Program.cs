// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Teams.Plugins.AspNetCore.Extensions;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api.Clients;
using System.Text.Json;

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
        await MentionUserAsync(context);
    else if (text.Contains("whoami"))
        await GetSingleMemberAsync(context);
    else
        await context.Send("Available commands:\n" +
                          "- **mention me** - Bot will mention you in the reply\n" +
                          "- **whoami** - Get your user information");
});

// Welcome new members to the conversation
teamsApp.OnConversationUpdate(async context =>
{
    if (context.Activity.MembersAdded != null)
    {
        foreach (var member in context.Activity.MembersAdded)
        {
            if (member.Id != context.Activity.Recipient.Id)
            {
                await context.Send($"Welcome to the team {member.Name}!\n\n" +
                                  "Welcome to Microsoft Teams conversationUpdate events demo bot.\n\n" +
                                  "Available commands:\n" +
                                  "- **mention me** - Bot will mention you in the reply\n" +
                                  "- **whoami** - Get your user information");
            }
        }
    }
});

webApp.Run();

// Mention user
async Task MentionUserAsync(IContext<MessageActivity> context)
{
    try
    {
        var userName = context.Activity.From.Name ?? "User";
        var mentionText = $"<at>{userName}</at>";        
        var messageActivity = new MessageActivity()
            .WithText($"Hello {mentionText}")
            .AddMention(context.Activity.From, addText: false);        
        await context.Send(messageActivity);
    }
    catch (Exception ex)
    {
        await context.Send($"Error sending mention: {ex.Message}");
    }
}

// Get and display user information
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
