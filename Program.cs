using System.Text.Json;
using System.Security.Principal;
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

        RunWindow(args is ["--adapters"]);
    }

    private static void RunWindow(bool forceAdapters)
    {
        using var identity = WindowsIdentity.GetCurrent();
        var administrator = new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
        string prefix = $@"Local\NetSwitchPro.{identity.User?.Value}";
        string elevated = prefix + ".elevated";
        if (!administrator && TryShow(elevated + ".show"))
            return;

        string instance = administrator ? elevated : prefix + ".normal";
        using var show = new EventWaitHandle(false, EventResetMode.AutoReset, instance + ".show");
        using var mutex = new Mutex(true, instance, out bool owns);
        if (!owns)
        {
            try { owns = mutex.WaitOne(0); }
            catch (AbandonedMutexException) { owns = true; }
        }
        if (!owns)
        {
            show.Set();
            return;
        }

        try
        {
            ApplicationConfiguration.Initialize();
            using var main = new MainForm(forceAdapters);
            var listener = ThreadPool.RegisterWaitForSingleObject(show, (_, _) =>
            {
                if (!main.IsHandleCreated || main.IsDisposed)
                    return;
                try { main.BeginInvoke(main.RestoreMainWindow); }
                catch (InvalidOperationException) { }
            }, null, Timeout.Infinite, executeOnlyOnce: false);
            try { Application.Run(main); }
            finally { listener.Unregister(null); }
        }
        finally { mutex.ReleaseMutex(); }
    }

    private static bool TryShow(string name)
    {
        try
        {
            if (!EventWaitHandle.TryOpenExisting(name, out var show))
                return false;
            using (show)
                return show.Set();
        }
        catch (UnauthorizedAccessException) { return false; }
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
