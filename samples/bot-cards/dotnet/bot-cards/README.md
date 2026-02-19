# Bot Cards - .NET (C#)

This sample demonstrates how to upload files and interact with adaptive cards in Microsoft Teams using a bot built with Teams SDK.

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)

## Run the sample

1. Navigate to this directory:
   ```bash
   cd dotnet/bot-cards
   ```

2. Copy the example launch settings file:
   ```bash
   cp Properties/launchSettings.EXAMPLE.json Properties/launchSettings.json
   ```
   
   Update `Properties/launchSettings.json` with your Teams app credentials (TenantId, ClientId, ClientSecret).

3. Restore dependencies and run:
   ```bash
   dotnet run
   ```

The bot will start listening on `http://localhost:3978`.

## Next Steps

Refer to the [main README](../../README.md) for instructions on how to:
- Deploy and test your bot in Microsoft Teams
- Configure Teams app manifest and credentials
