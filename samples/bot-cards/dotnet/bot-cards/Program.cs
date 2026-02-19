// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Teams.Plugins.AspNetCore.Extensions;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Invokes;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Samples.BotCards.Handlers;
using System.Net.Http.Headers;
using System.Text.Json;
using FileDownloadInfoModel = Microsoft.Teams.Samples.BotCards.Models.FileDownloadInfo;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddHttpClient();
builder.AddTeams();

var webApp = builder.Build();
var teamsApp = webApp.UseTeams(true);

var filesPath = Path.Combine(Environment.CurrentDirectory, "Files");
if (!Directory.Exists(filesPath))
{
    Directory.CreateDirectory(filesPath);
}

var httpClientFactory = webApp.Services.GetRequiredService<IHttpClientFactory>();

// Handle bot installation (send welcome message when bot is first added)
teamsApp.OnInstall(async context =>
{
    await SendWelcomeMessage(context);
});

// Handle incoming messages
teamsApp.OnMessage(async context =>
{
    var activity = context.Activity;
    var text = activity.Text?.Trim() ?? "";
    var attachment = activity.Attachments?.FirstOrDefault();

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

        // Handle card-related commands
        if (normalizedText.Contains("card actions"))
        {
            await Cards.SendAdaptiveCardActions(context);
            return;
        }
        else if (normalizedText.Contains("togglevisibility"))
        {
            await Cards.SendToggleVisibilityCard(context);
            return;
        }
        // Handle file commands
        else if (normalizedText.Contains("send file") || normalizedText == "file")
        {
            await Attachments.SendFileCard(context, filesPath);
            return;
        }
        // Unrecognized command - show welcome
        else
        {
            await SendWelcomeMessage(context);
            return;
        }
    }

    // Handle file attachments
    if (attachment != null)
    {
        var contentTypeValue = attachment.ContentType?.Value ?? attachment.ContentType?.ToString() ?? "";

        // Handle file downloads
        if (contentTypeValue == "application/vnd.microsoft.teams.file.download.info")
        {
            try
            {
                var fileDownloadInfo = attachment.Content != null 
                    ? JsonSerializer.Deserialize<FileDownloadInfoModel>((JsonElement)attachment.Content)
                    : null;

                if (fileDownloadInfo != null)
                {
                    var httpClient = httpClientFactory.CreateClient();
                    var response = await httpClient.GetAsync(fileDownloadInfo.DownloadUrl);
                    response.EnsureSuccessStatusCode();

                    var fileName = attachment.Name ?? $"download_{DateTime.Now.Ticks}";
                    var filePath = Path.Combine(filesPath, fileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        await response.Content.CopyToAsync(fileStream);
                    }

                    var successMessage = new MessageActivity().WithText($"File <b>{fileName}</b> downloaded successfully!").WithTextFormat(TextFormat.Xml);
                    await context.Send(successMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error downloading file: {ex}");
            }
        }
        // Handle inline images
        else if (contentTypeValue.StartsWith("image/"))
        {
            await Attachments.ProcessInlineImage(context, attachment, filesPath, httpClientFactory);
            return;
        }
        else
        {
            await Attachments.SendFileCard(context, filesPath);
            return;
        }
    }
    // No text or attachment - show welcome
    else
    {
        await SendWelcomeMessage(context);
    }
});

// Handle file consent responses
teamsApp.OnFileConsent(async context =>
{
    var fileConsentResponse = context.Activity.Value;

    if (fileConsentResponse != null && fileConsentResponse.Action == Microsoft.Teams.Api.Action.Accept)
    {
        try
        {
            var contextData = fileConsentResponse.Context != null
                ? JsonSerializer.Deserialize<Dictionary<string, string>>((JsonElement)fileConsentResponse.Context)
                : null;

            var fileName = contextData?["filename"] ?? "file.txt";
            var filePath = Path.Combine(filesPath, fileName);

            if (!File.Exists(filePath))
            {
                await context.Send($"File {fileName} not found.");
                return;
            }

            var fileData = await File.ReadAllBytesAsync(filePath);
            var uploadInfo = fileConsentResponse.UploadInfo;

            var httpClient = httpClientFactory.CreateClient();
            var fileContent = new ByteArrayContent(fileData);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            fileContent.Headers.ContentRange = new ContentRangeHeaderValue(0, fileData.Length - 1, fileData.Length);

            var uploadResponse = await httpClient.PutAsync(uploadInfo!.UploadUrl, fileContent);
            uploadResponse.EnsureSuccessStatusCode();

            // Send success message with download link and file info
            var downloadCard = new
            {
                uniqueId = uploadInfo.UniqueId,
                fileType = uploadInfo.FileType
            };

            var fileInfoAttachment = new Attachment
            {
                ContentType = new ContentType("application/vnd.microsoft.teams.card.file.info"),
                Name = fileName,
                ContentUrl = uploadInfo.ContentUrl,
                Content = downloadCard
            };

            var successMessage = new MessageActivity().WithText($"<b> File uploaded successfully!</b><br/><br/>Your file <b>{fileName}</b> has been uploaded to OneDrive.<br/><br/><a href=\"{uploadInfo.ContentUrl}\">Click here to open in OneDrive</a>").WithTextFormat(TextFormat.Xml);
            successMessage.Attachments = new List<Attachment> { fileInfoAttachment };
            await context.Send(successMessage);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error uploading file: {ex}");
        }
    }
    else if (fileConsentResponse != null && fileConsentResponse.Action == Microsoft.Teams.Api.Action.Decline)
    {
        try
        {
            var contextData = fileConsentResponse.Context != null
                ? JsonSerializer.Deserialize<Dictionary<string, string>>((JsonElement)fileConsentResponse.Context)
                : null;

            var fileName = contextData?["filename"] ?? "file";

            var declineMessage = new MessageActivity().WithText($"Declined. We won't upload file <b>{fileName}</b>.").WithTextFormat(TextFormat.Xml);
            await context.Send(declineMessage);
        }
        catch
        {
            Console.WriteLine("Error processing file decline response");
        }
    }
});

webApp.Run();

// Sends a welcome message
async Task SendWelcomeMessage<T>(IContext<T> context) where T : IActivity
{
    await context.Send("Welcome to the Teams Bot Cards!");
}
