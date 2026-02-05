/*
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
*/

using Microsoft.Teams.Plugins.AspNetCore.Extensions;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api.Clients;
using System.Text.Json;
using System.Collections.Concurrent;
using System.Text.RegularExpressions;

// Load environment variables (handled by ASP.NET Core configuration system)
var builder = WebApplication.CreateBuilder(args);
builder.AddTeams();
var webApp = builder.Build();
var configuration = webApp.Services.GetRequiredService<IConfiguration>();
var teamsApp = webApp.UseTeams(true);

// Simple in-memory storage for conversation references (for proactive messaging)
// In production, use persistent storage like a database
var conversation_storage = new ConcurrentDictionary<string, string>();

// Get and display information about the current user.
async Task get_member_info(IContext<MessageActivity> ctx)
{
    try
    {
        var conversation_id = ctx.Activity.Conversation.Id;
        var user_id = ctx.Activity.From.Id;

        // Try to get member details from the API
        try
        {
            var members = await ctx.Api.Conversations.Members.GetAsync(conversation_id);
            var member = members?.FirstOrDefault(m => m.Id == user_id);

            if (member != null)
            {
                await ctx.Send($"You are: {member.Name}");
            }
            else
            {
                await ctx.Send($"You are: {ctx.Activity.From.Name}");
            }
        }
        catch (Exception)
        {
            // Fallback to activity sender info
            await ctx.Send($"You are: {ctx.Activity.From.Name}");
        }
    }
    catch (Exception e)
    {
        if (!e.Message.Contains("MemberNotFoundInConversation"))
        {
            throw;
        }
    }
}

// Mention the current user in a text message.
async Task mention_user(IContext<MessageActivity> ctx)
{
    try
    {
        var conversation_id = ctx.Activity.Conversation.Id;
        var user_id = ctx.Activity.From.Id;

        // Try to get member details from the API
        var member = ctx.Activity.From;
        try
        {
            var members = await ctx.Api.Conversations.Members.GetAsync(conversation_id);
            var foundMember = members?.FirstOrDefault(m => m.Id == user_id);

            if (foundMember != null)
            {
                member = foundMember;
            }
        }
        catch (Exception)
        {
            member = ctx.Activity.From;
        }

        var member_name = member.Name ?? "User";
        var member_id = member.Id;

        // Create the mention text
        var mention_text = $"<at>{member_name}</at>";

        // Send message with mention entity
        var activity = new MessageActivity()
            .WithText($"Hello {mention_text}!")
            .AddMention(member, addText: false);

        await ctx.Send(activity);
    }
    catch (Exception e)
    {
        if (!e.Message.Contains("MemberNotFoundInConversation"))
        {
            throw;
        }
    }
}

// Remove bot mention from the text.
string remove_recipient_mention(string? text, string? recipient_name = null)
{
    if (string.IsNullOrEmpty(text))
    {
        return "";
    }

    // Remove <at>...</at> tags
    var cleaned = Regex.Replace(text, @"<at>.*?</at>", "");
    // Remove extra whitespace
    cleaned = string.Join(" ", cleaned.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
    return cleaned.Trim();
}

// Send a proactive message to a user.
async Task<bool> send_proactive_notification(string user_id, string message = "Hey! This is a proactive message from the bot!")
{
    if (!conversation_storage.TryGetValue(user_id, out var conversation_id) || string.IsNullOrEmpty(conversation_id))
    {
        return false;
    }

    var activity = new MessageActivity().WithText(message);
    await teamsApp.Send(conversation_id, activity);
    return true;
}

// Send a proactive message after a delay.
async Task delayed_proactive_message(string user_id, int delay_seconds = 10)
{
    await Task.Delay(TimeSpan.FromSeconds(delay_seconds));
    await send_proactive_notification(
        user_id,
        $"Reminder: This proactive message was sent {delay_seconds} seconds after your request!"
    );
}

// Handle incoming message activities.
teamsApp.OnMessage(async (IContext<MessageActivity> ctx) =>
{
    // Get the text and clean it (remove bot mention)
    var raw_text = ctx.Activity.Text ?? "";
    var recipient_name = ctx.Activity.Recipient?.Name;
    var text = remove_recipient_mention(raw_text, recipient_name).ToLower();

    // Store conversation reference for proactive messaging (from any message)
    var user_aad_id = ctx.Activity.From.AadObjectId;
    if (!string.IsNullOrEmpty(user_aad_id))
    {
        conversation_storage.AddOrUpdate(user_aad_id, ctx.Activity.Conversation.Id, (key, oldValue) => ctx.Activity.Conversation.Id);
    }

    // Handle proactive messaging command
    if (text.Contains("proactive"))
    {
        if (!string.IsNullOrEmpty(user_aad_id))
        {
            await ctx.Send("Got it! I'll send you a proactive message in 10 seconds...");
            // Schedule the proactive message (runs in background)
            _ = Task.Run(async () => await delayed_proactive_message(user_aad_id, 10));
        }
        else
        {
            await ctx.Send("Sorry, I couldn't identify your user ID for proactive messaging.");
        }
        return;
    }

    // Handle different commands
    if (text.Contains("mentionme"))
    {
        await mention_user(ctx);
        return;
    }

    if (text.Contains("who") || text.Contains("whoami"))
    {
        await get_member_info(ctx);
        return;
    }

    // Echo back any other message
    if (!string.IsNullOrEmpty(text))
    {
        await ctx.Send($"Echo: {text}");
    }
});

// Handle when new members are added to the conversation.
teamsApp.OnConversationUpdate(async (IContext<ConversationUpdateActivity> ctx) =>
{
    var members_added = ctx.Activity.MembersAdded;
    if (members_added == null || !members_added.Any())
    {
        return;
    }

    var recipient_id = ctx.Activity.Recipient?.Id;
    var conversation_type = ctx.Activity.Conversation?.Id?.Contains("personal");

    foreach (var member in members_added)
    {
        // If the bot itself was added, send welcome message
        if (member.Id == recipient_id)
        {
            var welcome_message =
                "Welcome to Microsoft Teams conversationUpdate events demo bot.\n\n" +
                "Available commands:\n" +
                "- **mention me** - Bot will mention you in the reply\n" +
                "- **whoami** - Get your user information\n" +
                "- **proactive** - Bot will send a proactive message in 10 seconds\n" +
                "- **Echo** - Bot will echo your message back";

            await ctx.Send(welcome_message);
        }
        // If another member was added to a non-personal conversation
        else if (conversation_type == false)
        {
            var given_name = "";
            var surname = "";
            var name = $"{given_name} {surname}".Trim();
            if (string.IsNullOrEmpty(name))
            {
                name = member.Name ?? "User";
            }
            await ctx.Send($"Welcome to the team {name}");
        }
    }
});

webApp.Run();
