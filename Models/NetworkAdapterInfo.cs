namespace NetworkAdapterSwitcher.Models;

internal sealed record NetworkAdapterInfo(
    Guid Id,
    string Name,
    string Description,
    string Type,
    bool Enabled,
    bool Connected,
    bool IsVirtual,
    bool IsBluetooth,
    string? IPv4 = null);
