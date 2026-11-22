using System.Net.Http;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace NetworkAdapterSwitcher.Services;

internal sealed record AvailableUpdate(Version Version, Uri Download);

internal static class ReleaseUpdates
{
    internal const string ReleasesPage = "https://github.com/timenti/NetSwitchPro/releases/latest";
    private const string Api = "https://api.github.com/repos/timenti/NetSwitchPro/releases/latest";
    private static readonly Regex TagFormat = new(@"\Av(?:0|[1-9][0-9]*)(?:\.(?:0|[1-9][0-9]*)){0,2}\z", RegexOptions.CultureInvariant);

    private static readonly HttpClient Client = new()
    {
        Timeout = TimeSpan.FromSeconds(20),
        MaxResponseContentBufferSize = 256 * 1024
    };

    public static async Task<AvailableUpdate?> CheckAsync(Version current, CancellationToken cancellation)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, Api);
        request.Headers.UserAgent.ParseAdd("NetSwitchPro/" + current.ToString(3));
        request.Headers.Accept.ParseAdd("application/vnd.github+json");
        request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");
        using var response = await Client.SendAsync(request, cancellation).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        return ReadRelease(await response.Content.ReadAsStringAsync(cancellation).ConfigureAwait(false), current);
    }

    internal static string DisplayVersion(Version version) =>
        version.Build > 0 ? version.ToString(3) :
        version.Minor > 0 ? version.ToString(2) :
        version.Major.ToString(CultureInfo.InvariantCulture);

    internal static AvailableUpdate? ReadRelease(string json, Version current)
    {
        var release = JsonSerializer.Deserialize<Release>(json) ?? throw new InvalidDataException("GitHub returned no release.");
        if (release.Draft || release.Prerelease)
            return null;
        if (release.Tag is not { } tag || !TagFormat.IsMatch(tag))
            throw new InvalidDataException("Unexpected release version from GitHub.");

        string numericTag = tag[1..];
        if (!Version.TryParse(numericTag.Contains('.') ? numericTag : numericTag + ".0", out var parsed))
            throw new InvalidDataException("Unexpected release version from GitHub.");

        var version = new Version(parsed.Major, parsed.Minor, Math.Max(0, parsed.Build), 0);
        if (version <= current)
            return null;

        string name = $"NetSwitchPro-{version.ToString(3)}.exe";
        if (release.Assets?.Any(asset => asset is { Size: > 0 } && asset.Name == name) != true)
            throw new InvalidDataException("The GitHub release has no NetSwitch Pro installer.");
        return new(version, new Uri($"https://github.com/timenti/NetSwitchPro/releases/download/{tag}/{name}"));
    }

    private sealed record Release(
        [property: JsonPropertyName("tag_name")] string? Tag,
        [property: JsonPropertyName("draft")] bool Draft,
        [property: JsonPropertyName("prerelease")] bool Prerelease,
        [property: JsonPropertyName("assets")] Asset[]? Assets);

    private sealed record Asset(
        [property: JsonPropertyName("name")] string Name,
        [property: JsonPropertyName("size")] long Size);
}
