using NetworkAdapterSwitcher.Models;

namespace NetworkAdapterSwitcher.Services;

internal sealed record ConnectionPair(NetworkEndpoint? First, NetworkEndpoint? Second)
{
    public bool HasBoth => First is not null && Second is not null;
    public bool UsesDifferentAdapters => HasBoth && First!.AdapterId != Second!.AdapterId;

    public bool Contains(NetworkEndpoint endpoint) => endpoint == First || endpoint == Second;

    public bool ContainsAdapter(Guid id) => First?.AdapterId == id || Second?.AdapterId == id;

    public NetworkEndpoint? ResolveTarget(
        IReadOnlyList<NetworkAdapterInfo> adapters, IReadOnlyList<WifiAdapter> wireless)
    {
        if (First is not { } first || Second is not { } second || first == second)
            return null;

        var a = adapters.FirstOrDefault(adapter => adapter.Id == first.AdapterId);
        if (a is null)
            return null;

        if (first.AdapterId == second.AdapterId)
        {
            if (!a.IsWifi || first.WifiProfile is null || second.WifiProfile is null)
                return null;

            var active = wireless.FirstOrDefault(adapter => adapter.Id == a.Id);
            return WifiService.ChooseTarget(first.WifiProfile, second.WifiProfile, active?.ActiveProfile) == first.WifiProfile
                ? first : second;
        }

        var b = adapters.FirstOrDefault(adapter => adapter.Id == second.AdapterId);
        if (b is null)
            return null;
        if (a.Enabled && b.Enabled && a.Preferred == b.Preferred)
        {
            bool aReady = IsConnected(first, adapters, wireless);
            bool bReady = IsConnected(second, adapters, wireless);
            if (aReady != bReady)
                return aReady ? second : first;
        }

        return AdapterSwitch.ChooseTarget(a, b).Id == a.Id ? first : second;
    }

    public Guid? OtherAdapterFor(NetworkEndpoint target) =>
        UsesDifferentAdapters && ContainsAdapter(target.AdapterId)
            ? (First!.AdapterId == target.AdapterId ? Second!.AdapterId : First.AdapterId)
            : null;

    public NetworkEndpoint? Current(
        IReadOnlyList<NetworkAdapterInfo> adapters, IReadOnlyList<WifiAdapter> wireless)
    {
        var target = ResolveTarget(adapters, wireless);
        var source = target is null ? null : OtherAdapterFor(target);
        var primary = adapters.FirstOrDefault(adapter => adapter.Id == source && adapter.Connected)
            ?? adapters.FirstOrDefault(adapter => ContainsAdapter(adapter.Id) && adapter.Preferred && adapter.Connected)
            ?? adapters.FirstOrDefault(adapter => ContainsAdapter(adapter.Id) && adapter.Connected)
            ?? adapters.FirstOrDefault(adapter => adapter.Preferred && adapter.Connected)
            ?? adapters.FirstOrDefault(adapter => adapter.Connected);
        if (primary is null)
            return null;

        string? profile = primary.IsWifi
            ? wireless.FirstOrDefault(adapter => adapter.Id == primary.Id)?.ActiveProfile
            : null;
        return new NetworkEndpoint(primary.Id, profile);
    }

    public static bool Matches(NetworkEndpoint? current, NetworkEndpoint choice) =>
        current?.AdapterId == choice.AdapterId && (choice.WifiProfile is null || choice.WifiProfile == current.WifiProfile);

    private static bool IsConnected(NetworkEndpoint endpoint,
        IReadOnlyList<NetworkAdapterInfo> adapters, IReadOnlyList<WifiAdapter> wireless) =>
        endpoint.WifiProfile is null
            ? adapters.Any(adapter => adapter.Id == endpoint.AdapterId && adapter.Connected)
            : wireless.Any(adapter => adapter.Id == endpoint.AdapterId && adapter.Connected && adapter.ActiveProfile == endpoint.WifiProfile);
}
