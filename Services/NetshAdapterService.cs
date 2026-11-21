using System.Globalization;
using System.Management;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using NetworkAdapterSwitcher.Models;

namespace NetworkAdapterSwitcher.Services;

internal interface IAdapterService
{
    IReadOnlyList<NetworkAdapterInfo> GetAdapters();
    void SetEnabled(Guid id, bool enabled);
}

internal sealed class WindowsAdapterService : IAdapterService
{
    private const string ModernScope = @"root\StandardCimv2";
    private static readonly string[] VirtualTerms = ["virtual", "wireguard", "wintun", "radmin", "tailscale", "vpn", "loopback"];

    public IReadOnlyList<NetworkAdapterInfo> GetAdapters()
    {
        var links = ReadWindowsLinks();
        var interfaces = ReadInterfaces();
        uint preferred = GetBestInterface(0x01010101, out uint index) == 0 ? index : 0;
        using var search = new ManagementObjectSearcher(
            "SELECT GUID, NetConnectionID, Name, AdapterTypeID, InterfaceIndex, NetConnectionStatus, NetEnabled, " +
            "ConfigManagerErrorCode, PhysicalAdapter, PNPDeviceID, ServiceName FROM Win32_NetworkAdapter");
        using var rows = search.Get();
        var adapters = new Dictionary<Guid, NetworkAdapterInfo>();

        foreach (ManagementObject row in rows)
        {
            using (row)
            {
                if (!Guid.TryParse(row["GUID"]?.ToString(), out var id) ||
                    row["NetConnectionID"] is not string name || string.IsNullOrWhiteSpace(name))
                    continue;

                uint error = Number(row["ConfigManagerErrorCode"]);
                links.TryGetValue(id, out var link);
                if (error is 24 or 45 || link?.Operational == 6)
                    continue;

                interfaces.TryGetValue(id, out var nic);
                var description = row["Name"]?.ToString() ?? name;
                var pnp = row["PNPDeviceID"]?.ToString() ?? "";
                var service = row["ServiceName"]?.ToString() ?? "";
                var identity = $"{name} {description} {pnp} {service}".ToLowerInvariant();
                uint typeId = Number(row["AdapterTypeID"]);
                bool bluetooth = identity.Contains("bluetooth") || identity.Contains("btpan") ||
                                 pnp.StartsWith(@"BTH\", StringComparison.OrdinalIgnoreCase);
                bool hardwareBus = pnp.StartsWith(@"PCI\", StringComparison.OrdinalIgnoreCase) ||
                                   pnp.StartsWith(@"USB\", StringComparison.OrdinalIgnoreCase);
                bool virtualAdapter = !bluetooth && !hardwareBus && link?.Hardware != true &&
                    (row["PhysicalAdapter"] is false ||
                     pnp.StartsWith(@"ROOT\", StringComparison.OrdinalIgnoreCase) ||
                     pnp.StartsWith(@"SWD\", StringComparison.OrdinalIgnoreCase) || VirtualTerms.Any(identity.Contains));
                bool wifi = link?.Type == 71 || nic?.Type == NetworkInterfaceType.Wireless80211 ||
                            typeId == 9 || identity.Contains("wi-fi") || identity.Contains("wireless");
                string type = bluetooth ? "Bluetooth" : wifi ? "Wi-Fi" : virtualAdapter ? "Virtual" :
                    link?.Type == 6 || typeId is 0 or 5 ? "Ethernet" : "Network";
                var connection = row["NetConnectionStatus"] as ushort?;
                var (enabled, connected) = ResolveLink(link, nic, error, row["NetEnabled"] as bool?, connection);
                uint interfaceIndex = link?.Index ?? Number(row["InterfaceIndex"]);
                adapters.TryAdd(id, new(id, name, description, type, enabled, connected, virtualAdapter && !wifi,
                    bluetooth, nic?.IPv4, preferred != 0 && interfaceIndex == preferred));
            }
        }

        foreach (var (id, link) in links)
        {
            if (adapters.ContainsKey(id) || link.Hidden || link.Operational == 6 || string.IsNullOrEmpty(link.Name))
                continue;
            interfaces.TryGetValue(id, out var nic);
            bool wifi = link.Type == 71 || nic?.Type == NetworkInterfaceType.Wireless80211;
            bool bluetooth = link.Description.Contains("bluetooth", StringComparison.OrdinalIgnoreCase);
            bool virtualAdapter = !link.Hardware && !wifi && !bluetooth;
            string type = bluetooth ? "Bluetooth" : wifi ? "Wi-Fi" : virtualAdapter ? "Virtual" :
                link.Type == 6 ? "Ethernet" : "Network";
            var (enabled, connected) = ResolveLink(link, nic, 0, null, null);
            adapters.TryAdd(id, new(id, link.Name, link.Description, type, enabled, connected, virtualAdapter,
                bluetooth, nic?.IPv4, preferred != 0 && link.Index == preferred));
        }

        return adapters.Values.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase).ToArray();
    }

    private static (bool Enabled, bool Connected) ResolveLink(
        WindowsLink? link, InterfaceDetails? nic, uint error, bool? legacyEnabled, ushort? legacyConnection)
    {
        bool connected = link is { } current
            ? current.Media == 1 || (current.Media == 0 && (nic?.Up == true || current.Operational == 1))
            : nic?.Up == true || legacyConnection == 2;
        bool enabled = link?.Admin switch
        {
            1 => true,
            2 or 3 => false,
            _ => error != 22 && legacyConnection != 5 &&
                (connected || legacyConnection is 0 or 7 || (legacyEnabled ?? legacyConnection.HasValue))
        };
        return (enabled, enabled && connected);
    }

    private static Dictionary<Guid, WindowsLink> ReadWindowsLinks()
    {
        var links = new Dictionary<Guid, WindowsLink>();
        try
        {
            using var search = new ManagementObjectSearcher(ModernScope,
                "SELECT Name, InterfaceGuid, InterfaceDescription, InterfaceIndex, InterfaceAdminStatus, " +
                "InterfaceOperationalStatus, MediaConnectState, InterfaceType, HardwareInterface, Hidden FROM MSFT_NetAdapter");
            using var rows = search.Get();
            foreach (ManagementObject row in rows)
            {
                using (row)
                {
                    if (Guid.TryParse(row["InterfaceGuid"]?.ToString(), out var id))
                        links.TryAdd(id, new(row["Name"]?.ToString() ?? "", row["InterfaceDescription"]?.ToString() ?? "",
                            Number(row["InterfaceIndex"]), Number(row["InterfaceAdminStatus"]), Number(row["InterfaceOperationalStatus"]),
                            Number(row["MediaConnectState"]), Number(row["InterfaceType"]), row["HardwareInterface"] is true,
                            row["Hidden"] is true));
                }
            }
        }
        catch (Exception ex) when (ex is ManagementException or UnauthorizedAccessException or COMException) { }
        return links;
    }

    private static Dictionary<Guid, InterfaceDetails> ReadInterfaces()
    {
        var interfaces = new Dictionary<Guid, InterfaceDetails>();
        try
        {
            foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (!Guid.TryParse(nic.Id, out var id))
                    continue;
                string? address = null;
                try
                {
                    address = nic.GetIPProperties().UnicastAddresses
                        .FirstOrDefault(a => a.Address.AddressFamily == AddressFamily.InterNetwork)?.Address.ToString();
                }
                catch (NetworkInformationException) { }
                interfaces.TryAdd(id, new(nic.NetworkInterfaceType, nic.OperationalStatus == OperationalStatus.Up, address));
            }
        }
        catch (NetworkInformationException) { }
        return interfaces;
    }

    public void SetEnabled(Guid id, bool enabled)
    {
        // A typed GUID cannot match a different adapter with the same display name.
        using var search = new ManagementObjectSearcher($"SELECT * FROM Win32_NetworkAdapter WHERE GUID='{id:B}'");
        using var rows = search.Get();
        using var adapter = rows.Cast<ManagementObject>().SingleOrDefault();
        if (adapter is not null)
        {
            InvokeChange(adapter, enabled);
            return;
        }

        using var modern = new ManagementObjectSearcher(ModernScope, $"SELECT * FROM MSFT_NetAdapter WHERE InterfaceGuid='{id:B}'");
        using var modernRows = modern.Get();
        using var item = modernRows.Cast<ManagementObject>().SingleOrDefault()
            ?? throw new InvalidOperationException($"Adapter unavailable ({id:B}).");
        InvokeChange(item, enabled);
    }

    private static void InvokeChange(ManagementObject adapter, bool enabled)
    {
        object? result = adapter.InvokeMethod(enabled ? "Enable" : "Disable", null);
        if (result is null || Number(result) != 0)
            throw new InvalidOperationException($"Windows did not change adapter state. WMI returned {result ?? "no result"}.");
    }

    private static uint Number(object? value) => Convert.ToUInt32(value, CultureInfo.InvariantCulture);

    private sealed record WindowsLink(string Name, string Description, uint Index, uint Admin,
        uint Operational, uint Media, uint Type, bool Hardware, bool Hidden);
    private sealed record InterfaceDetails(NetworkInterfaceType Type, bool Up, string? IPv4);

    [DllImport("iphlpapi.dll")]
    private static extern uint GetBestInterface(uint destinationAddress, out uint bestInterfaceIndex);
}
