// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Teams.Plugins.AspNetCore.Extensions;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Samples.BotCards.Handlers;
using System.Text.Json;

var builder = WebApplication.CreateBuilder(args);
builder.AddTeams();

var webApp = builder.Build();
var teamsApp = webApp.UseTeams(true);

// Handle incoming messages
teamsApp.OnMessage(async context =>
{
    var activity = context.Activity;
    var text = activity.Text?.Trim() ?? "";

    // Handle data submission from adaptive cards (activity.Value)
    if (activity.Value != null)
    {
        var dataSubmitted = JsonSerializer.Serialize(activity.Value);
        Console.WriteLine($"Data submitted: {dataSubmitted}");
        await context.Send($"Data Submitted: {dataSubmitted}");
        return;
    }

    // Handle text commands
    if (!string.IsNullOrEmpty(text))
    {
        var normalizedText = text.ToLower();

        if (normalizedText.Contains("card actions"))
        {
            await Cards.SendAdaptiveCardActions(context);
            return;
        }
        else if (normalizedText.Contains("toggle visibility"))
        {
            await Cards.SendToggleVisibilityCard(context);
            return;
        }
    }

    // Default - show welcome message
    await SendWelcomeMessage(context);
});

webApp.Run();

// Sends a welcome message
async Task SendWelcomeMessage<T>(IContext<T> context) where T : IActivity
{
    await context.Send("Welcome to the Cards Bot! To interact with me, send one of the following commands: 'card actions' or 'toggle visibility'");
}
