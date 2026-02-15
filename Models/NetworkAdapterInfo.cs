namespace NetworkAdapterSwitcher.Models;

internal sealed record NetworkAdapterInfo(
    string Name,
    string FullName,
    string Type,
    bool IsAdminEnabled,
    bool IsConnected,
    string? IPv4Address,
    bool IsVirtual,
    bool IsBluetooth)
{
    public string StatusText => IsAdminEnabled ? "Enabled" : "Disabled";
}
