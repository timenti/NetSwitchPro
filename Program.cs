using System.Text.Json;
using System.Diagnostics;
using System.ComponentModel;
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

        bool installUpdate = args is [ReleaseUpdates.InstallArgument];
        if (installUpdate && !UpdatePublisher.IsTrusted(Application.ExecutablePath))
            return;

        RunWindow(args is [StartupRegistration.TrayArgument], installUpdate);
    }

    private static void RunWindow(bool startInTray, bool installUpdate)
    {
        using var identity = WindowsIdentity.GetCurrent();
        bool administrator = new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
        string prefix = $@"Local\NetSwitchPro.{identity.User?.Value}";
        string ownSlot = prefix + (administrator ? ".elevated" : ".normal");
        string otherSlot = prefix + (administrator ? ".normal" : ".elevated");
        var version = typeof(Program).Assembly.GetName().Version!;
        using var launch = new Mutex(false, prefix + ".launch");
        if (!RunningInstance.TryEnter(launch, TimeSpan.FromSeconds(8)))
        {
            if (!startInTray && !installUpdate)
                ShowLaunchMessage("NetSwitch Pro is finishing another launch.", "Другой запуск NetSwitch Pro ещё выполняется.",
                    MessageBoxButtons.OK);
            return;
        }

        Mutex? ownership = null;
        RunningInstance? instance = null;
        Exception? setupError = null;
        try
        {
            // A normal launch can show the elevated copy. An elevated launch replaces
            // the normal copy of the same version so changing access still works.
            if (HandledByPeer(otherSlot, version, startInTray, installUpdate, replaceEqual: administrator && !startInTray) ||
                HandledByPeer(ownSlot, version, startInTray, installUpdate, replaceEqual: false))
                return;

            ownership = new Mutex(false, ownSlot);
            while (!RunningInstance.TryEnter(ownership, TimeSpan.Zero))
            {
                if (HandledByPeer(ownSlot, version, startInTray, installUpdate, replaceEqual: false))
                    return;
            }

            if (installUpdate)
            {
                try
                {
                    var installed = StartupRegistration.InstallCurrentCopy(Application.ExecutablePath);
                    using var child = Process.Start(new ProcessStartInfo(installed) { UseShellExecute = false });
                    if (child is not null)
                    {
                        ownership.ReleaseMutex();
                        return;
                    }
                }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException or Win32Exception)
                {
                    setupError = error;
                }
            }

            instance = new RunningInstance(ownSlot, version);
        }
        finally
        {
            launch.ReleaseMutex();
            if (instance is null)
                ownership?.Dispose();
        }

        if (instance is null || ownership is null)
            return;

        try
        {
            ApplicationConfiguration.Initialize();
            using var context = new ApplicationContext();
            using var main = new MainForm();
            if (setupError is not null)
                main.ReportStartupError(setupError);
            try { StartupRegistration.UpdateForNewerVersion(Application.ExecutablePath); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException or System.Security.SecurityException)
            {
                main.ReportStartupError(error);
            }

            void Queue(Action action)
            {
                if (!main.IsHandleCreated || main.IsDisposed)
                    return;
                try { main.BeginInvoke(action); }
                catch (InvalidOperationException) { }
            }

            main.FormClosed += (_, _) => context.ExitThread();
            _ = main.Handle;
            var showListener = ThreadPool.RegisterWaitForSingleObject(instance.Show, (_, _) => Queue(main.RestoreMainWindow),
                null, Timeout.Infinite, executeOnlyOnce: false);
            var retireListener = ThreadPool.RegisterWaitForSingleObject(instance.Retire, (_, _) => Queue(main.RequestUpgradeExit),
                null, Timeout.Infinite, executeOnlyOnce: true);
            try
            {
                if (startInTray)
                    main.BeginInvoke(main.InitializeTray);
                else
                    main.Show();
                Application.Run(context);
            }
            finally
            {
                retireListener.Unregister(null);
                showListener.Unregister(null);
            }
        }
        finally
        {
            instance.Dispose();
            ownership.ReleaseMutex();
            ownership.Dispose();
        }
    }

    private static bool HandledByPeer(string name, Version version, bool startInTray, bool quiet, bool replaceEqual)
    {
        while (!RunningInstance.WaitUntilFree(name, TimeSpan.Zero))
        {
            var other = RunningInstance.ReadVersion(name);
            if (other is not null && (other > version || (other == version && !replaceEqual)))
            {
                if (startInTray || RunningInstance.Signal(name, ".show"))
                    return true;
                if (RunningInstance.WaitUntilFree(name, TimeSpan.Zero))
                    return false;
                if (!quiet)
                    ShowLaunchMessage("NetSwitch Pro is already running. Open it from the notification area.",
                    "NetSwitch Pro уже запущен. Откройте его из области уведомлений.", MessageBoxButtons.OK);
                return true;
            }

            if (other is not null && RunningInstance.Signal(name, ".retire") &&
                RunningInstance.WaitUntilFree(name, TimeSpan.FromSeconds(quiet ? 70 : 12)))
                return false;

            if (startInTray)
                return true;

            if (RunningInstance.WaitUntilFree(name, TimeSpan.Zero))
                return false;
            if (quiet)
                return true;
            var decision = ShowLaunchMessage(
                "Another version is still running. Choose Exit from its tray menu, then click Retry.",
                "В трее работает другая версия. Закройте её через «Выход» в меню значка, затем нажмите кнопку повтора.",
                MessageBoxButtons.RetryCancel);
            if (decision != DialogResult.Retry)
                return true;
        }
        return false;
    }

    private static DialogResult ShowLaunchMessage(string english, string russian, MessageBoxButtons buttons) =>
        MessageBox.Show(System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "ru" ? russian : english,
            "NetSwitch Pro", buttons, MessageBoxIcon.Information);

    private static void WriteReadOnlyReport(string report)
    {
        var result = new Dictionary<string, object>();
        using var wifi = new WifiService();
        try
        {
            var interfaces = wifi.GetAdapters();
            result["wifiInterfaces"] = interfaces.Count;
            result["savedProfiles"] = interfaces.Sum(i => i.Profiles.Length);
            result["connectedInterfaces"] = interfaces.Count(i => i.Connected);
            result["activeProfilesRecognized"] = interfaces.Count(i => i.ActiveProfile is not null);
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
