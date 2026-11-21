namespace NetworkAdapterSwitcher.Models;

internal sealed record NetworkEndpoint(Guid AdapterId, string? WifiProfile = null);
