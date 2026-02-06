# Bot File Upload - .NET (C#)

This sample demonstrates a File Upload Bot for Microsoft Teams using C#. The bot can send files to users and receive files from users.

## Features

- Send files to users: Say "send file" to receive a file from the bot
- Receive files from users: Upload files to the bot and it will save them
- Handle inline images: Send images directly in chat
- File consent flow: Proper consent handling for file uploads

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)

## Run the sample

1. Navigate to this directory:
   ```bash
   cd dotnet/bot-file-upload
   ```

2. Configure your `appsettings.json` file with your credentials:
   ```json
   {
     "TENANT_ID": "your-tenant-id",
     "CLIENT_ID": "your-client-id",
     "CLIENT_SECRET": "your-client-secret"
   }
   ```

3. Restore dependencies and run:
   ```bash
   dotnet run
   ```

The bot will start listening on `http://localhost:3978`.

## Bot Commands

| Command | Description |
| --- | --- |
| send file or file | Bot sends a file consent card for downloading a file |
| Send any file | Bot downloads and saves the file |
| Send an image | Bot saves the image and echoes it back |
| Any other message | Bot echoes the message back |

## Further Reading

- [File Upload in Teams](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bots-filesv4)
- [Microsoft Teams SDK Documentation](https://learn.microsoft.com/microsoftteams/platform/)