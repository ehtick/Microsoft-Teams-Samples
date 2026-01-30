// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Teams.Plugins.AspNetCore.Extensions;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Invokes;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Serialization;

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

// Handle incoming messages
teamsApp.OnMessage(async context =>
{
    var activity = context.Activity;

    // Check if message contains actual file attachments
    bool hasFileAttachment = activity.Attachments != null &&
                             activity.Attachments.Count > 0 &&
                             activity.Attachments[0].ContentType?.Value != "text/html";

    if (hasFileAttachment)
    {
        var attachment = activity.Attachments![0];
        var contentTypeValue = attachment.ContentType?.Value ?? attachment.ContentType?.ToString() ?? "";

        // Handle file downloads
        if (contentTypeValue == "application/vnd.microsoft.teams.file.download.info")
        {
            try
            {
                var fileDownloadInfo = JsonSerializer.Deserialize<FileDownloadInfo>(
                    JsonSerializer.Serialize(attachment.Content));

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

                    var successMessage = new MessageActivity($"File <b>{fileName}</b> downloaded successfully!");
                    successMessage.TextFormat = TextFormat.Xml;
                    await context.Send(successMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error downloading file: {ex}");
                await context.Send("Sorry, there was an error downloading the file. Please try again later.");
            }
        }
        // Handle inline images
        else if (contentTypeValue.StartsWith("image/"))
        {
            try
            {
                var httpClient = httpClientFactory.CreateClient();
                var response = await httpClient.GetAsync(attachment.ContentUrl);
                response.EnsureSuccessStatusCode();

                var fileName = $"ImageFromUser_{DateTime.Now.Ticks}.png";
                var filePath = Path.Combine(filesPath, fileName);

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await response.Content.CopyToAsync(fileStream);
                }

                var imageData = Convert.ToBase64String(File.ReadAllBytes(filePath));
                var inlineAttachment = new Attachment
                {
                    Name = fileName,
                    ContentType = new ContentType("image/png"),
                    ContentUrl = $"data:image/png;base64,{imageData}"
                };

                var replyMessage = new MessageActivity($"Received and saved your image. File size: {response.Content.Headers.ContentLength} bytes");
                replyMessage.Attachments = new List<Attachment> { inlineAttachment };
                await context.Send(replyMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing image: {ex}");
                await context.Send("Sorry, there was an error processing your image. Please try again later.");
            }
        }
        else
        {
            await context.Send($"File attachment received but type '{contentTypeValue}' not supported for processing.");
        }
    }
    else
    {
        // Send a file consent card
        try
        {
            var fileName = "teams-logo.png";
            var filePath = Path.Combine(filesPath, fileName);

            if (!File.Exists(filePath))
            {
                await File.WriteAllTextAsync(filePath, "Sample file content for Teams file upload demo.");
            }

            var fileInfo = new FileInfo(filePath);
            var fileSize = fileInfo.Length;

            var fileConsentCard = new FileConsentCard
            {
                Name = fileName,
                Description = "This is the file I want to send you",
                SizeInBytes = fileSize,
                AcceptContext = new { fileName = fileName },
                DeclineContext = new { fileName = fileName }
            };

            var consentAttachment = new Attachment
            {
                ContentType = new ContentType(FileConsentCard.ContentType),
                Name = fileName,
                Content = fileConsentCard
            };

            var message = new MessageActivity("Please accept the file");
            message.Attachments = new List<Attachment> { consentAttachment };
            await context.Send(message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating file consent card: {ex}");
            await context.Send("Sorry, there was an error preparing the file. Please try again later.");
        }
    }
});

// Handle file consent responses
teamsApp.OnFileConsent(async context =>
{
    var fileConsentResponse = context.Activity.Value;

    if (fileConsentResponse?.Action == "accept")
    {
        try
        {
            var contextData = JsonSerializer.Deserialize<Dictionary<string, string>>(
                JsonSerializer.Serialize(fileConsentResponse.Context));

            var fileName = contextData?["fileName"] ?? "file.txt";
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

            var lowerFileName = fileName.ToLower();
            if (lowerFileName.EndsWith(".png") || lowerFileName.EndsWith(".jpg") ||
                lowerFileName.EndsWith(".jpeg") || lowerFileName.EndsWith(".gif"))
            {
                var imageData = Convert.ToBase64String(fileData);
                var mimeType = lowerFileName.EndsWith(".png") ? "image/png" :
                              lowerFileName.EndsWith(".gif") ? "image/gif" : "image/jpeg";

                var imageAttachment = new Attachment
                {
                    Name = fileName,
                    ContentType = new ContentType(mimeType),
                    ContentUrl = $"data:{mimeType};base64,{imageData}"
                };

                var successMessage = new MessageActivity($"<b>File uploaded successfully.</b> Your file <b>{fileName}</b> has been uploaded to OneDrive.");
                successMessage.TextFormat = TextFormat.Xml;
                successMessage.Attachments = new List<Attachment> { imageAttachment };
                await context.Send(successMessage);
            }
            else
            {
                var successMessage = new MessageActivity($"<b>File uploaded successfully.</b> Your file <b>{fileName}</b> has been uploaded to <a href=\"{uploadInfo.ContentUrl}\">OneDrive</a>. Click the link to view or download.");
                successMessage.TextFormat = TextFormat.Xml;
                await context.Send(successMessage);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error uploading file: {ex}");
            await context.Send("Sorry, there was an error uploading the file. Please try again later.");
        }
    }
    else if (fileConsentResponse?.Action == "decline")
    {
        try
        {
            var contextData = JsonSerializer.Deserialize<Dictionary<string, string>>(
                JsonSerializer.Serialize(fileConsentResponse.Context));

            var fileName = contextData?["fileName"] ?? "file";

            var declineMessage = new MessageActivity($"Declined. We won't upload file <b>{fileName}</b>.");
            declineMessage.TextFormat = TextFormat.Xml;
            await context.Send(declineMessage);
        }
        catch
        {
            await context.Send("You declined the file upload.");
        }
    }
});

webApp.Run();

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