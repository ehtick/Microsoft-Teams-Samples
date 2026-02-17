// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api;
using System.Text.Json.Serialization;

namespace Microsoft.Teams.Samples.BotCards.Handlers;

public static class Attachments
{
    // Send file consent card
    public static async Task SendFileCard(dynamic context, string filesPath)
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
                SizeInBytes = fileSize,
                Name = filename,
                AcceptContext = consentContext,
                DeclineContext = consentContext
            };

            var message = new MessageActivity();
            message.Attachments = new List<Attachment>
            {
                new Attachment
                {
                    Content = fileCard,
                    ContentType = new ContentType("application/vnd.microsoft.teams.card.file.consent"),
                    Name = filename
                }
            };
            await context.Send(message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending file card: {ex}");
            await context.Send($"Error sending file card: {ex.Message}");
        }
    }

    // Process inline image
    public static async Task ProcessInlineImage(dynamic context, Attachment file, string filesPath, IHttpClientFactory httpClientFactory)
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

            var message = new MessageActivity($"Image <b>{fileName}</b> of size <b>{fileSize}</b> bytes received and saved.");
            message.Attachments = new List<Attachment> { inlineAttachment };
            await context.Send(message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing inline image: {ex}");
            await context.Send($"Error processing image: {ex.Message}");
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

// Model classes
public class FileConsentCard
{
    public const string ContentType = "application/vnd.microsoft.teams.card.file.consent";

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;

    [JsonPropertyName("sizeInBytes")]
    public long SizeInBytes { get; set; }

    [JsonPropertyName("acceptContext")]
    public object? AcceptContext { get; set; }

    [JsonPropertyName("declineContext")]
    public object? DeclineContext { get; set; }
}

public class FileDownloadInfo
{
    [JsonPropertyName("downloadUrl")]
    public string DownloadUrl { get; set; } = string.Empty;

    [JsonPropertyName("uniqueId")]
    public string? UniqueId { get; set; }

    [JsonPropertyName("fileType")]
    public string? FileType { get; set; }
}
