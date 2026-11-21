using System.Text.Json;
using NetworkAdapterSwitcher.Services;

namespace NetworkAdapterSwitcher;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        if (args is ["--probe", var report])
        {
            WriteReadOnlyReport(report);
            return;
        }

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm(args is ["--adapters"]));
    }

    private static void WriteReadOnlyReport(string report)
    {
        var result = new Dictionary<string, object>();
        using var wifi = new WifiService();
        try
        {
            var interfaces = wifi.GetAdapters(readCurrent: false);
            result["wifiInterfaces"] = interfaces.Count;
            result["savedProfiles"] = interfaces.Sum(i => i.Profiles.Length);
            result["connectedInterfaces"] = interfaces.Count(i => i.Connected);
        }
        catch (Exception ex) { result["wifiError"] = ex.Message; }

        try
        {
            var adapters = new WindowsAdapterService().GetAdapters();
            result["adapters"] = adapters.Count;
            result["connectedAdapters"] = adapters.Count(a => a.Connected);
        }
        catch (Exception ex) { result["adapterError"] = ex.Message; }

        File.WriteAllText(report, JsonSerializer.Serialize(result));
    }
}
