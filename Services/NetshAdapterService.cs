using System.Globalization;
using System.Management;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using NetworkAdapterSwitcher.Models;

namespace NetworkAdapterSwitcher.Services;

internal interface IAdapterService
{
    IReadOnlyList<NetworkAdapterInfo> GetAdapters();
    void SetEnabled(Guid id, bool enabled);
}

internal sealed class WindowsAdapterService : IAdapterService
{
    public IReadOnlyList<NetworkAdapterInfo> GetAdapters()
    {
        using var search = new ManagementObjectSearcher(
            "SELECT GUID, NetConnectionID, Name, AdapterTypeID, NetConnectionStatus, NetEnabled, " +
            "ConfigManagerErrorCode, PhysicalAdapter, PNPDeviceID, ServiceName FROM Win32_NetworkAdapter");
        using var rows = search.Get();
        var addresses = ReadAddresses();
        var adapters = new Dictionary<Guid, NetworkAdapterInfo>();
        foreach (ManagementObject row in rows)
        {
            using (row)
            {
                if (!Guid.TryParse(row["GUID"]?.ToString(), out var id) ||
                    row["NetConnectionID"] is not string name || string.IsNullOrWhiteSpace(name))
                    continue;

                var error = row["ConfigManagerErrorCode"] as uint?;
                if (error is 24 or 45)
                    continue;

                var description = row["Name"]?.ToString() ?? name;
                var pnp = row["PNPDeviceID"]?.ToString() ?? "";
                var service = row["ServiceName"]?.ToString() ?? "";
                var identity = $"{name} {description} {pnp} {service}".ToLowerInvariant();
                var typeId = row["AdapterTypeID"] as ushort?;
                bool bluetooth = identity.Contains("bluetooth") || identity.Contains("btpan") ||
                                 pnp.StartsWith(@"BTH\", StringComparison.OrdinalIgnoreCase);
                bool hardwareBus = pnp.StartsWith(@"PCI\", StringComparison.OrdinalIgnoreCase) ||
                                   pnp.StartsWith(@"USB\", StringComparison.OrdinalIgnoreCase);
                bool virtualAdapter = !bluetooth && !hardwareBus &&
                    (row["PhysicalAdapter"] is false ||
                     pnp.StartsWith(@"ROOT\", StringComparison.OrdinalIgnoreCase) ||
                     pnp.StartsWith(@"SWD\", StringComparison.OrdinalIgnoreCase) ||
                     new[] { "virtual", "wireguard", "wintun", "radmin", "tailscale", "vpn", "loopback" }.Any(identity.Contains));
                string type = bluetooth ? "Bluetooth" : virtualAdapter ? "Virtual" :
                    typeId == 9 || identity.Contains("wi-fi") || identity.Contains("wireless") ? "Wi-Fi" :
                    typeId is 0 or 5 ? "Ethernet" : "Network";
                var connection = row["NetConnectionStatus"] as ushort?;
                bool enabled = error != 22 && (row["NetEnabled"] as bool? ?? connection != 5);
                adapters.TryAdd(id, new(id, name, description, type, enabled,
                    enabled && connection == 2, virtualAdapter, bluetooth, addresses.GetValueOrDefault(id)));
            }
        }

        return adapters.Values.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase).ToArray();
    }

    private static Dictionary<Guid, string> ReadAddresses()
    {
        var addresses = new Dictionary<Guid, string>();
        try
        {
            foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (Guid.TryParse(nic.Id, out var id) &&
                    nic.GetIPProperties().UnicastAddresses.FirstOrDefault(a =>
                        a.Address.AddressFamily == AddressFamily.InterNetwork) is { } address)
                    addresses.TryAdd(id, address.Address.ToString());
            }
        }
        catch (NetworkInformationException) { }
        return addresses;
    }

    public void SetEnabled(Guid id, bool enabled)
    {
        // A typed GUID cannot inject WQL syntax or match another adapter with the same name.
        using var search = new ManagementObjectSearcher(
            $"SELECT * FROM Win32_NetworkAdapter WHERE GUID='{id:B}'");
        using var rows = search.Get();
        using var adapter = rows.Cast<ManagementObject>().SingleOrDefault()
            ?? throw new InvalidOperationException($"Adapter unavailable ({id:B}).");
        object? result = adapter.InvokeMethod(enabled ? "Enable" : "Disable", null);
        if (result is null || Convert.ToUInt32(result, CultureInfo.InvariantCulture) != 0)
            throw new InvalidOperationException($"{adapter["NetConnectionID"]}: WMI returned {result ?? "no result"}.");
    }
}
