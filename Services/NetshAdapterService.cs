using System.Diagnostics;
using System.Management;
using System.Text;
using System.Text.Json;
using NetworkAdapterSwitcher.Models;

namespace NetworkAdapterSwitcher.Services;

internal sealed class NetshAdapterService
{
    public IReadOnlyList<NetworkAdapterInfo> GetAdapters()
    {
        var ipMap = LoadIpv4ByInterfaceIndex();
        var metadataByName = LoadWmiMetadataByConnectionName();

        var adapters = TryLoadFromPowerShell(ipMap, metadataByName);
        if (adapters.Count == 0)
        {
            adapters = LoadFromWmi(ipMap);
        }

        return adapters
            .GroupBy(a => a.Name, StringComparer.CurrentCultureIgnoreCase)
            .Select(g => g.First())
            .OrderBy(a => a.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    public void SetAdapterState(string adapterName, bool enable)
    {
        var escapedName = adapterName.Replace("'", "''", StringComparison.Ordinal);
        using var searcher = new ManagementObjectSearcher(
            $"SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID='{escapedName}' OR Name='{escapedName}'");

        using var adapters = searcher.Get();
        var adapter = adapters.Cast<ManagementObject>().FirstOrDefault()
            ?? throw new InvalidOperationException($"Network adapter '{adapterName}' was not found.");

        var method = enable ? "Enable" : "Disable";
        var result = adapter.InvokeMethod(method, null);
        var code = Convert.ToInt32(result);

        if (code is 0 or 1)
        {
            return;
        }

        throw new InvalidOperationException($"Failed to {(enable ? "enable" : "disable")} '{adapterName}'. WMI code: {code}.");
    }

    private static List<NetworkAdapterInfo> TryLoadFromPowerShell(Dictionary<uint, string> ipMap, Dictionary<string, WmiAdapterMetadata> metadataByName)
    {
        const string script = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8; Get-NetAdapter | Select-Object Name, InterfaceDescription, Status, ifIndex, ifType, InterfaceName, PnPDeviceID, Virtual, HardwareInterface, DriverDescription | ConvertTo-Json -Compress";
        var output = RunPowerShell(script);
        if (string.IsNullOrWhiteSpace(output))
        {
            return new List<NetworkAdapterInfo>();
        }

        try
        {
            using var doc = JsonDocument.Parse(output);
            var rows = doc.RootElement.ValueKind == JsonValueKind.Array
                ? doc.RootElement.EnumerateArray().ToList()
                : new List<JsonElement> { doc.RootElement };

            var adapters = new List<NetworkAdapterInfo>();
            foreach (var row in rows)
            {
                var name = GetString(row, "Name");
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var fullName = GetString(row, "InterfaceDescription");
                var status = GetString(row, "Status");
                var index = GetUInt(row, "ifIndex") ?? 0;
                var ifType = GetUInt(row, "ifType");
                var hardwareInterface = GetBool(row, "HardwareInterface");
                var isReportedVirtual = GetBool(row, "Virtual");
                var pnpDeviceId = GetString(row, "PnPDeviceID");
                var serviceName = GetString(row, "InterfaceName");
                var driverDescription = GetString(row, "DriverDescription");

                metadataByName.TryGetValue(name, out var metadata);
                ipMap.TryGetValue(index, out var ipv4);

                var isEnabled = !string.Equals(status, "Disabled", StringComparison.OrdinalIgnoreCase);
                var isConnected = string.Equals(status, "Up", StringComparison.OrdinalIgnoreCase);

                var resolvedFullName = string.IsNullOrWhiteSpace(fullName)
                    ? metadata?.FullName ?? name
                    : fullName;

                var effectiveIfType = ifType ?? metadata?.AdapterTypeId;
                var effectivePnp = string.IsNullOrWhiteSpace(pnpDeviceId) ? metadata?.PnpDeviceId : pnpDeviceId;
                var effectiveService = string.IsNullOrWhiteSpace(serviceName) ? metadata?.ServiceName : serviceName;
                var effectiveManufacturer = metadata?.Manufacturer;

                var isBluetooth = IsBluetoothAdapter(name, resolvedFullName, metadata?.Type, effectivePnp, effectiveService, effectiveManufacturer, effectiveIfType);
                var isVirtual = IsVirtualAdapter(
                    name,
                    resolvedFullName,
                    metadata?.Type,
                    metadata?.PhysicalAdapter,
                    effectiveIfType,
                    isReportedVirtual,
                    hardwareInterface,
                    effectivePnp,
                    effectiveService,
                    driverDescription,
                    isBluetooth);
                var type = ClassifyAdapterType(name, resolvedFullName, metadata?.Type, effectiveIfType, isBluetooth, isVirtual);

                adapters.Add(new NetworkAdapterInfo(
                    name,
                    resolvedFullName,
                    type,
                    isEnabled,
                    isConnected,
                    ipv4,
                    isVirtual,
                    isBluetooth));
            }

            return adapters;
        }
        catch
        {
            return new List<NetworkAdapterInfo>();
        }
    }

    private static List<NetworkAdapterInfo> LoadFromWmi(Dictionary<uint, string> ipMap)
    {
        var adapters = new List<NetworkAdapterInfo>();

        using var searcher = new ManagementObjectSearcher(
            "SELECT Index, NetConnectionID, Name, AdapterType, AdapterTypeID, NetConnectionStatus, ConfigManagerErrorCode, PhysicalAdapter, PNPDeviceID, ServiceName, Manufacturer FROM Win32_NetworkAdapter");

        foreach (ManagementObject adapter in searcher.Get())
        {
            var connectionName = adapter["NetConnectionID"]?.ToString();
            var fullName = adapter["Name"]?.ToString() ?? connectionName ?? string.Empty;
            if (string.IsNullOrWhiteSpace(connectionName))
            {
                continue;
            }

            var index = adapter["Index"] is uint idx ? idx : 0;
            var connectionStatus = adapter["NetConnectionStatus"] as ushort?;
            var physicalAdapter = adapter["PhysicalAdapter"] as bool?;
            var configError = adapter["ConfigManagerErrorCode"] as uint?;
            var adapterTypeId = adapter["AdapterTypeID"] as ushort?;
            var pnpDeviceId = adapter["PNPDeviceID"]?.ToString();
            var serviceName = adapter["ServiceName"]?.ToString();
            var manufacturer = adapter["Manufacturer"]?.ToString();
            var isDisabledByAdmin = configError == 22;
            var isEnabled = !isDisabledByAdmin;
            var isConnected = connectionStatus == 2;

            ipMap.TryGetValue(index, out var ipv4);

            var isBluetooth = IsBluetoothAdapter(connectionName, fullName, adapter["AdapterType"]?.ToString(), pnpDeviceId, serviceName, manufacturer, adapterTypeId);
            var isVirtual = IsVirtualAdapter(
                connectionName,
                fullName,
                adapter["AdapterType"]?.ToString(),
                physicalAdapter,
                adapterTypeId,
                null,
                null,
                pnpDeviceId,
                serviceName,
                null,
                isBluetooth);
            var type = ClassifyAdapterType(connectionName, fullName, adapter["AdapterType"]?.ToString(), adapterTypeId, isBluetooth, isVirtual);

            adapters.Add(new NetworkAdapterInfo(
                connectionName,
                fullName,
                type,
                isEnabled,
                isConnected,
                ipv4,
                isVirtual,
                isBluetooth));
        }

        return adapters;
    }

    private static Dictionary<string, WmiAdapterMetadata> LoadWmiMetadataByConnectionName()
    {
        var map = new Dictionary<string, WmiAdapterMetadata>(StringComparer.OrdinalIgnoreCase);

        using var searcher = new ManagementObjectSearcher(
            "SELECT NetConnectionID, Name, AdapterType, AdapterTypeID, PhysicalAdapter, PNPDeviceID, ServiceName, Manufacturer FROM Win32_NetworkAdapter");

        foreach (ManagementObject adapter in searcher.Get())
        {
            var connectionName = adapter["NetConnectionID"]?.ToString();
            if (string.IsNullOrWhiteSpace(connectionName) || map.ContainsKey(connectionName))
            {
                continue;
            }

            var fullName = adapter["Name"]?.ToString() ?? connectionName;
            map[connectionName] = new WmiAdapterMetadata(
                fullName,
                adapter["AdapterType"]?.ToString(),
                adapter["PhysicalAdapter"] as bool?,
                adapter["AdapterTypeID"] as ushort?,
                adapter["PNPDeviceID"]?.ToString(),
                adapter["ServiceName"]?.ToString(),
                adapter["Manufacturer"]?.ToString());
        }

        return map;
    }

    private static string? RunPowerShell(string script)
    {
        foreach (var shell in new[] { "powershell.exe", "pwsh" })
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = shell,
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{script}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(psi);
                if (process is null)
                {
                    continue;
                }

                var output = process.StandardOutput.ReadToEnd();
                _ = process.StandardError.ReadToEnd();
                process.WaitForExit(5000);

                if (process.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
                {
                    return output.Trim();
                }
            }
            catch
            {
                // fallback to next shell / WMI
            }
        }

        return null;
    }

    private static bool IsVirtualAdapter(
        string displayName,
        string fullName,
        string? sourceType,
        bool? physicalAdapter,
        uint? interfaceType,
        bool? reportedVirtual,
        bool? hardwareInterface,
        string? pnpDeviceId,
        string? serviceName,
        string? driverDescription,
        bool isBluetooth)
    {
        if (reportedVirtual == true)
        {
            return true;
        }

        var identity = NormalizeForMatch($"{displayName} {fullName} {sourceType} {serviceName} {driverDescription}");
        var pnp = NormalizeForMatch(pnpDeviceId ?? string.Empty);

        var hasVirtualMarkers = ContainsAny(identity,
            "virtual",
            "vethernet",
            "hyper-v",
            "hyperv",
            "vmware",
            "vbox",
            "loopback",
            "tunnel",
            "tap",
            "wintun",
            "wireguard",
            "zerotier",
            "tailscale",
            "hamachi",
            "openvpn",
            "protonvpn",
            "nordlynx",
            "cloudflarewarp",
            "radmin",
            "radminvpn");

        var pnpLooksVirtual = ContainsAny(pnp,
            "root\\",
            "swd\\",
            "vms_mp",
            "tap",
            "wintun",
            "wireguard",
            "radminvpn",
            "zerotier");

        var pnpLooksPhysical = pnp.StartsWith("pci\\", StringComparison.Ordinal) ||
                               pnp.StartsWith("usb\\", StringComparison.Ordinal) ||
                               pnp.StartsWith("pcip\\", StringComparison.Ordinal);

        if (pnpLooksPhysical)
        {
            return false;
        }

        if (interfaceType == 6)
        {
            return false;
        }

        if (isBluetooth)
        {
            return false;
        }

        if (interfaceType is 24 or 53 or 131)
        {
            return true;
        }

        if (hasVirtualMarkers || pnpLooksVirtual)
        {
            return true;
        }

        if (hardwareInterface == true && physicalAdapter != false)
        {
            return false;
        }

        if (physicalAdapter == true && hardwareInterface != false)
        {
            return false;
        }

        if (hardwareInterface == false && physicalAdapter == false)
        {
            return true;
        }

        if (physicalAdapter == false && !isBluetooth)
        {
            return true;
        }

        if (hardwareInterface == false && !isBluetooth)
        {
            return true;
        }

        return false;
    }

    private static bool IsBluetoothAdapter(string displayName, string fullName, string? sourceType, string? pnpDeviceId, string? serviceName, string? manufacturer, uint? interfaceType)
    {
        var text = NormalizeForMatch($"{displayName} {fullName} {sourceType} {serviceName} {manufacturer}");
        var pnp = NormalizeForMatch(pnpDeviceId ?? string.Empty);

        return ContainsAny(text, "bluetooth", "bth", "btpan") ||
               pnp.StartsWith("bth\\", StringComparison.Ordinal) ||
               interfaceType == 7;
    }

    private static string ClassifyAdapterType(string displayName, string fullName, string? sourceType, uint? interfaceType, bool isBluetooth, bool isVirtual)
    {
        var normalized = NormalizeForMatch($"{displayName} {fullName} {sourceType}");

        if (isBluetooth)
        {
            return "Bluetooth";
        }

        if (normalized.Contains("wi-fi") || normalized.Contains("wifi") || normalized.Contains("wireless") || interfaceType == 71)
        {
            return "Wi-Fi";
        }

        if (normalized.Contains("loopback") || interfaceType == 24)
        {
            return "Loopback";
        }

        if (normalized.Contains("tunnel") || interfaceType == 131)
        {
            return "Tunnel";
        }

        if (interfaceType == 23)
        {
            return "PPP";
        }

        if (interfaceType == 6 || normalized.Contains("ethernet"))
        {
            return "Ethernet";
        }

        if (isVirtual)
        {
            return "Virtual";
        }

        return string.IsNullOrWhiteSpace(sourceType) ? "Unknown" : sourceType;
    }

    private static Dictionary<uint, string> LoadIpv4ByInterfaceIndex()
    {
        var map = new Dictionary<uint, string>();

        using var searcher = new ManagementObjectSearcher(
            "SELECT InterfaceIndex, IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True");

        foreach (ManagementObject config in searcher.Get())
        {
            if (config["InterfaceIndex"] is not uint index)
            {
                continue;
            }

            var addresses = config["IPAddress"] as string[];
            var ipv4 = addresses?.FirstOrDefault(ip => ip.Contains('.', StringComparison.Ordinal));
            if (!string.IsNullOrWhiteSpace(ipv4))
            {
                map[index] = ipv4;
            }
        }

        return map;
    }

    private static string NormalizeForMatch(string source)
    {
        return source
            .ToLowerInvariant()
            .Replace('–', '-')
            .Replace('—', '-')
            .Replace('‑', '-')
            .Replace(" ", string.Empty);
    }

    private static bool ContainsAny(string source, params string[] markers)
    {
        foreach (var marker in markers)
        {
            if (source.Contains(marker, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static string? GetString(JsonElement row, string name)
    {
        return row.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null
            ? value.GetString()
            : null;
    }

    private static uint? GetUInt(JsonElement row, string name)
    {
        if (!row.TryGetProperty(name, out var value))
        {
            return null;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var i) && i >= 0)
        {
            return (uint)i;
        }

        if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out var parsed) && parsed >= 0)
        {
            return (uint)parsed;
        }

        return null;
    }

    private static bool? GetBool(JsonElement row, string name)
    {
        if (!row.TryGetProperty(name, out var value))
        {
            return null;
        }

        if (value.ValueKind == JsonValueKind.True)
        {
            return true;
        }

        if (value.ValueKind == JsonValueKind.False)
        {
            return false;
        }

        return null;
    }

    private sealed record WmiAdapterMetadata(
        string FullName,
        string? Type,
        bool? PhysicalAdapter,
        ushort? AdapterTypeId,
        string? PnpDeviceId,
        string? ServiceName,
        string? Manufacturer);
}
