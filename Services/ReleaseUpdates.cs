using System.Diagnostics;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NetworkAdapterSwitcher.Services;

internal sealed record AvailableUpdate(Version Version, Uri Download, long Length, byte[] Digest);

internal static class ReleaseUpdates
{
    internal const string ReleasesPage = "https://github.com/timenti/NetSwitchPro/releases/latest";
    internal const string InstallArgument = "--install-update";
    private const string Api = "https://api.github.com/repos/timenti/NetSwitchPro/releases/latest";
    private const long MaximumBytes = 128 * 1024 * 1024;

    private static readonly HttpClient Client = new()
    {
        Timeout = TimeSpan.FromSeconds(20),
        MaxResponseContentBufferSize = 256 * 1024
    };

    public static async Task<AvailableUpdate?> CheckAsync(Version current, CancellationToken cancellation)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, Api);
        request.Headers.UserAgent.ParseAdd("NetSwitchPro/2.1");
        request.Headers.Accept.ParseAdd("application/vnd.github+json");
        request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");
        using var response = await Client.SendAsync(request, cancellation).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        var data = await response.Content.ReadAsStringAsync(cancellation).ConfigureAwait(false);
        return ReadRelease(data, current);
    }

    internal static AvailableUpdate? ReadRelease(string json, Version current)
    {
        var release = JsonSerializer.Deserialize<Release>(json) ?? throw new InvalidDataException("GitHub returned no release.");
        if (release.Draft || release.Prerelease)
            return null;
        if (release.Tag is not { Length: > 1 } tag || tag[0] != 'v' ||
            !Version.TryParse(tag[1..], out var parsed) || parsed.Build < 0 || parsed.Revision != -1 ||
            tag != "v" + parsed.ToString(3))
            throw new InvalidDataException("Unexpected release version from GitHub.");

        var version = new Version(parsed.Major, parsed.Minor, parsed.Build, 0);
        if (version <= current)
            return null;

        string name = $"NetSwitchPro-{parsed.ToString(3)}.exe";
        var asset = release.Assets?.FirstOrDefault(item => item.Name == name);
        if (asset is not { Size: > 0 and <= MaximumBytes, Digest.Length: 71 } ||
            !asset.Digest.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The GitHub release has no verifiable Windows installer.");
        byte[] digest;
        try { digest = Convert.FromHexString(asset.Digest[7..]); }
        catch (FormatException error) { throw new InvalidDataException("The release checksum is invalid.", error); }

        return new(version, new Uri($"https://github.com/timenti/NetSwitchPro/releases/download/{tag}/{name}"), asset.Size, digest);
    }

    public static async Task<string> DownloadAsync(AvailableUpdate update, IProgress<int> progress, CancellationToken cancellation)
    {
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        deadline.CancelAfter(TimeSpan.FromMinutes(4));
        using var request = new HttpRequestMessage(HttpMethod.Get, update.Download);
        request.Headers.UserAgent.ParseAdd("NetSwitchPro/2.1");
        using var response = await Client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token)
            .ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        if (response.Content.Headers.ContentLength is { } length && length != update.Length)
            throw new InvalidDataException("The download size does not match the release.");

        string folder = Path.Combine(StartupRegistration.ProgramDirectory, "Updates");
        Directory.CreateDirectory(folder);
        string name = $"NetSwitchPro-{update.Version.ToString(3)}-{Guid.NewGuid():N}";
        string temporary = Path.Combine(folder, name + ".tmp");
        string target = Path.Combine(folder, name + ".exe");
        try
        {
            await using var input = await response.Content.ReadAsStreamAsync(deadline.Token).ConfigureAwait(false);
            await SaveVerifiedAsync(input, temporary, update, progress, deadline.Token).ConfigureAwait(false);
            var file = FileVersionInfo.GetVersionInfo(temporary);
            if (!StartupRegistration.IsNetSwitch(file) || StartupRegistration.VersionOf(file) != update.Version)
                throw new InvalidDataException("The downloaded executable reports an unexpected product or version.");
            if (!UpdatePublisher.MatchesInstalledPublisher(Environment.ProcessPath!, temporary))
                throw new InvalidDataException("The update is missing a valid signature from the installed publisher.");
            File.Move(temporary, target);
            return target;
        }
        finally
        {
            if (File.Exists(temporary))
                File.Delete(temporary);
        }
    }

    internal static async Task SaveVerifiedAsync(Stream input, string file, AvailableUpdate update,
        IProgress<int> progress, CancellationToken cancellation)
    {
        await using var output = new FileStream(file, FileMode.CreateNew, FileAccess.Write, FileShare.None, 65536, useAsync: true);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[65536];
        long total = 0;
        int reported = -1;
        int count;
        while ((count = await input.ReadAsync(buffer, cancellation).ConfigureAwait(false)) > 0)
        {
            total += count;
            if (total > update.Length)
                throw new InvalidDataException("The download exceeded its declared size.");
            hash.AppendData(buffer, 0, count);
            await output.WriteAsync(buffer.AsMemory(0, count), cancellation).ConfigureAwait(false);
            int percent = (int)(total * 100 / update.Length);
            if (percent != reported)
            {
                reported = percent;
                progress.Report(percent);
            }
        }

        if (total != update.Length || !CryptographicOperations.FixedTimeEquals(hash.GetHashAndReset(), update.Digest))
            throw new InvalidDataException("The download failed its SHA-256 integrity check.");
    }

    private sealed record Release(
        [property: JsonPropertyName("tag_name")] string? Tag,
        [property: JsonPropertyName("draft")] bool Draft,
        [property: JsonPropertyName("prerelease")] bool Prerelease,
        [property: JsonPropertyName("assets")] Asset[]? Assets);

    private sealed record Asset(
        [property: JsonPropertyName("name")] string Name,
        [property: JsonPropertyName("size")] long Size,
        [property: JsonPropertyName("digest")] string? Digest);
}
