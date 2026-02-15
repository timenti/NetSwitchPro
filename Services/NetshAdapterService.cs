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
        const string script = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8; Get-NetAdapter | Select-Object Name, InterfaceDescription, Status, ifIndex | ConvertTo-Json -Compress";
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
                var name = row.TryGetProperty("Name", out var n) ? n.GetString() : null;
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var fullName = row.TryGetProperty("InterfaceDescription", out var d) ? d.GetString() : null;
                var status = row.TryGetProperty("Status", out var s) ? s.GetString() : null;
                var index = row.TryGetProperty("ifIndex", out var i) && i.TryGetInt32(out var i32) && i32 >= 0
                    ? (uint)i32
                    : 0;

                metadataByName.TryGetValue(name, out var metadata);
                ipMap.TryGetValue(index, out var ipv4);

                var isEnabled = !string.Equals(status, "Disabled", StringComparison.OrdinalIgnoreCase);
                var isConnected = string.Equals(status, "Up", StringComparison.OrdinalIgnoreCase);

                var resolvedFullName = string.IsNullOrWhiteSpace(fullName)
                    ? metadata?.FullName ?? name
                    : fullName;
                var type = metadata?.Type ?? "Unknown";
                var physicalAdapter = metadata?.PhysicalAdapter;

                adapters.Add(new NetworkAdapterInfo(
                    name,
                    resolvedFullName,
                    type,
                    isEnabled,
                    isConnected,
                    ipv4,
                    IsVirtualAdapter(name, resolvedFullName, type, physicalAdapter),
                    IsBluetoothAdapter(name, resolvedFullName, type)));
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
            "SELECT Index, NetConnectionID, Name, AdapterType, NetConnectionStatus, ConfigManagerErrorCode, PhysicalAdapter FROM Win32_NetworkAdapter");

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
            var isDisabledByAdmin = configError == 22;
            var enabled = !isDisabledByAdmin;
            var isConnected = connectionStatus == 2;
            var type = adapter["AdapterType"]?.ToString();
            if (string.IsNullOrWhiteSpace(type))
            {
                type = "Unknown";
            }

            ipMap.TryGetValue(index, out var ipv4);

            adapters.Add(new NetworkAdapterInfo(
                connectionName,
                fullName,
                type,
                enabled,
                isConnected,
                ipv4,
                IsVirtualAdapter(connectionName, fullName, type, physicalAdapter),
                IsBluetoothAdapter(connectionName, fullName, type)));
        }

        return adapters;
    }

    private static Dictionary<string, WmiAdapterMetadata> LoadWmiMetadataByConnectionName()
    {
        var map = new Dictionary<string, WmiAdapterMetadata>(StringComparer.OrdinalIgnoreCase);

        using var searcher = new ManagementObjectSearcher(
            "SELECT NetConnectionID, Name, AdapterType, PhysicalAdapter FROM Win32_NetworkAdapter");

        foreach (ManagementObject adapter in searcher.Get())
        {
            var connectionName = adapter["NetConnectionID"]?.ToString();
            if (string.IsNullOrWhiteSpace(connectionName) || map.ContainsKey(connectionName))
            {
                continue;
            }

            var fullName = adapter["Name"]?.ToString() ?? connectionName;
            var type = adapter["AdapterType"]?.ToString();
            if (string.IsNullOrWhiteSpace(type))
            {
                type = "Unknown";
            }

            map[connectionName] = new WmiAdapterMetadata(fullName, type, adapter["PhysicalAdapter"] as bool?);
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

    private static bool IsVirtualAdapter(string displayName, string fullName, string type, bool? physicalAdapter)
    {
        if (physicalAdapter == true)
        {
            return false;
        }

        if (physicalAdapter == false)
        {
            return true;
        }

        var text = NormalizeForMatch(displayName + " " + fullName + " " + type);
        return text.Contains("virtual") ||
               text.Contains("vmware") ||
               text.Contains("vbox") ||
               text.Contains("hyper-v") ||
               text.Contains("hyperv") ||
               text.Contains("vethernet") ||
               text.Contains("virtualethernet");
    }

    private static bool IsBluetoothAdapter(string displayName, string fullName, string type)
    {
        var text = NormalizeForMatch(displayName + " " + fullName + " " + type);
        return text.Contains("bluetooth");
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

    private sealed record WmiAdapterMetadata(string FullName, string Type, bool? PhysicalAdapter);
}
