// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api;
using Microsoft.Teams.Apps;
using System.Text.Json.Serialization;
using FileDownloadInfoModel = Microsoft.Teams.Samples.BotCards.Models.FileDownloadInfo;

namespace Microsoft.Teams.Samples.BotCards.Handlers;

public static class Attachments
{
    // Send file consent card
    public static async Task SendFileCard<T>(IContext<T> context, string filesPath) where T : IActivity
    {
        try
        {
            var filename = "teams-logo.png";
            var filePath = Path.Combine(filesPath, filename);
            
            // Check if teams logo exists
            if (!File.Exists(filePath))
            {
                await context.Send($"Error: {filename} not found in Files folder. Please add the teams-logo.png file.");
                return;
            }

            var stats = new FileInfo(filePath);
            var fileSize = stats.Length;
            var consentContext = new { filename = filename };

            var fileCard = new FileConsentCard
            {
                Description = "This is the file I want to send you",
                SizeInBytes = (int)fileSize,
                AcceptContext = consentContext,
                DeclineContext = consentContext
            };

            var message = new MessageActivity
            {
                Attachments = new List<Attachment>
                {
                    new Attachment
                    {
                        Content = fileCard,
                        ContentType = new ContentType("application/vnd.microsoft.teams.card.file.consent"),
                        Name = filename
                    }
                }
            };
            await context.Send(message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending file card: {ex}");
        }
    }

    // Process inline image
    public static async Task ProcessInlineImage<T>(IContext<T> context, Attachment file, string filesPath, IHttpClientFactory httpClientFactory) where T : IActivity
    {
        try
        {
            var httpClient = httpClientFactory.CreateClient();
            var response = await httpClient.GetAsync(file.ContentUrl);
            response.EnsureSuccessStatusCode();

            var fileName = await GenerateFileName(filesPath);
            var filePath = Path.Combine(filesPath, fileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await response.Content.CopyToAsync(fileStream);
            }

            var fileSize = new FileInfo(filePath).Length;
            var inlineAttachment = GetInlineAttachment(fileName, filesPath);

            var message = new MessageActivity($"Image <b>{fileName}</b> of size <b>{fileSize}</b> bytes received and saved.")
            {
                Attachments = new List<Attachment> { inlineAttachment }
            };
            await context.Send(message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing inline image: {ex}");
        }
    }

    // Get inline attachment
    public static Attachment GetInlineAttachment(string fileName, string filesPath)
    {
        var imageData = File.ReadAllBytes(Path.Combine(filesPath, fileName));
        var base64Image = Convert.ToBase64String(imageData);
        return new Attachment
        {
            Name = fileName,
            ContentType = new ContentType("image/png"),
            ContentUrl = $"data:image/png;base64,{base64Image}"
        };
    }

    // Generate sequential file name
    public static async Task<string> GenerateFileName(string fileDir)
    {
        const string filenamePrefix = "UserAttachment";
        var files = Directory.GetFiles(fileDir);
        var filteredFiles = files
            .Select(f => Path.GetFileName(f))
            .Where(f => f.Contains(filenamePrefix))
            .Select(f =>
            {
                var parts = f.Split(filenamePrefix);
                if (parts.Length > 1)
                {
                    var numStr = parts[1].Split('.')[0];
                    return int.TryParse(numStr, out var num) ? num : 0;
                }
                return 0;
            })
            .Where(num => num > 0)
            .ToList();

        var maxSeq = filteredFiles.Any() ? filteredFiles.Max() : 0;
        return $"{filenamePrefix}{maxSeq + 1}.png";
    }
}
