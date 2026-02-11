// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Teams.Plugins.AspNetCore.Extensions;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api.Clients;
using System.Collections.Concurrent;

// Initialize Teams App - automatically uses CLIENT_ID and CLIENT_SECRET from environment variables
// Note: .env file is only required when running on Teams (not needed for local development with devtools)
var builder = WebApplication.CreateBuilder(args);
builder.AddTeams();
var webApp = builder.Build();
var teamsApp = webApp.UseTeams(true);

// Simple in-memory storage for conversation references (for proactive messaging)
// In production, use persistent storage like a database
var conversationStorage = new ConcurrentDictionary<string, string>();

// Send a proactive message to a user
async Task<bool> SendProactiveNotification(
    string userId,
    string message = "Hey! This is a proactive message from the bot!")
{
    if (!conversationStorage.TryGetValue(userId, out var conversationId) || string.IsNullOrEmpty(conversationId))
    {
        return false;
    }

    await teamsApp.Send(conversationId, new MessageActivity().WithText(message));
    return true;
}

// Send a proactive message after a delay
async Task DelayedProactiveMessage(string userId, int delaySeconds = 10)
{
    await Task.Delay(TimeSpan.FromSeconds(delaySeconds));
    await SendProactiveNotification(
        userId,
        $"Reminder: This proactive message was sent {delaySeconds} seconds after your request!"
    );
}

// Handle conversation update events (when bot is added or members join)
teamsApp.OnConversationUpdate(async (IContext<ConversationUpdateActivity> context) =>
{
    var membersAdded = context.Activity.MembersAdded;
    if (membersAdded != null)
    {
        foreach (var member in membersAdded)
        {
            // Check if bot was added to the conversation
            if (member.Id == context.Activity.Recipient?.Id)
            {
                await SendWelcomeMessage(context);
            }
        }
    }
});

// Handles incoming messages and routes to appropriate functions based on message content
teamsApp.OnMessage(async (IContext<MessageActivity> context) =>
{
    // Get message text and normalize it
    var text = (context.Activity.Text ?? "").Trim().ToLower();

    // Store conversation reference for proactive messaging (from any message)
    var userAadId = context.Activity.From.AadObjectId;
    if (!string.IsNullOrEmpty(userAadId))
    {
        conversationStorage.AddOrUpdate(userAadId, context.Activity.Conversation.Id, (key, oldValue) => context.Activity.Conversation.Id);
    }

    // Handle proactive messaging command
    if (text.Contains("proactive"))
    {
        if (!string.IsNullOrEmpty(userAadId))
        {
            await context.Send("Got it! I'll send you a proactive message in 10 seconds...");
            // Schedule the proactive message (runs in background)
            _ = Task.Run(async () =>
            {
                try
                {
                    await DelayedProactiveMessage(userAadId, 10);
                }
                catch (Exception err)
                {
                    Console.WriteLine($"Error sending proactive message: {err.Message}");
                }
            });
        }
        else
        {
            await context.Send("Sorry, I couldn't identify your user ID for proactive messaging.");
        }
        return;
    }

    // Handle mention me command
    if (text.Contains("mentionme") || text.Contains("mention me"))
    {
        await MentionUser(context);
    }
    // Handle whoami command
    else if (text.Contains("whoami"))
    {
        await GetSingleMember(context);
    }
    // Handle welcome command
    else if (text.Contains("welcome"))
    {
        await SendWelcomeMessage(context);
    }
    // Echo greeting messages
    else if (text.Contains("hi") || text.Contains("hello"))
    {
        await EchoMessage(context, text);
    }
    else
    {
        await SendWelcomeMessage(context);
    }
});

// Sends a welcome message
async Task SendWelcomeMessage<T>(IContext<T> context) where T : IActivity
{
    await context.Send("Welcome to the Teams Quickstart Bot!");
}

// Echo back the user's message
async Task EchoMessage(IContext<MessageActivity> context, string text)
{
    await context.Send($"**Echo :** {text}");
}

// Retrieves and displays information about the current user
async Task GetSingleMember(IContext<MessageActivity> context)
{
    var conversationId = context.Activity.Conversation.Id;
    var userId = context.Activity.From.Id;

    try
    {
        var members = await context.Api.Conversations.Members.GetAsync(conversationId);
        var member = members?.FirstOrDefault(m => m.Id == userId);

        if (member != null)
        {
            await context.Send($"You are: {member.Name}");
        }
    }
    catch (Exception error)
    {
        Console.WriteLine($"Error getting member: {error.Message}");
    }
}

// Mention a user in a message
async Task MentionUser(IContext<MessageActivity> context)
{
    var conversationId = context.Activity.Conversation.Id;
    var userId = context.Activity.From.Id;

    try
    {
        var members = await context.Api.Conversations.Members.GetAsync(conversationId);
        var member = members?.FirstOrDefault(m => m.Id == userId);

        if (member != null)
        {
            // Create a text message with user mention
            var mentionText = $"<at>{member.Name}</at>";
            var activity = new MessageActivity()
                .WithText($"Hello {mentionText}")
                .AddMention(member, addText: false);

            await context.Send(activity);
        }
    }
    catch (Exception error)
    {
        Console.WriteLine($"Error mentioning user: {error.Message}");
    }
}

// Starts the Teams bot application and listens for incoming requests
webApp.Run();
