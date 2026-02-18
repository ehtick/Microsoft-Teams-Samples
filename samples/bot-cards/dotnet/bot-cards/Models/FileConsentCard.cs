// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json.Serialization;

namespace Microsoft.Teams.Samples.BotCards.Models;

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
