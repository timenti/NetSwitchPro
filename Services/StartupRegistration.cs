using System.Diagnostics;
using Microsoft.Win32;

namespace NetworkAdapterSwitcher.Services;

internal static class StartupRegistration
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    internal const string EntryName = "NetSwitchPro";
    internal const string TrayArgument = "--start-in-tray";
    private const string CommandSuffix = "\" " + TrayArgument;

    public static string? ReadCommand()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKey);
        return key?.GetValue(EntryName) as string;
    }

    public static string CommandFor(string executable)
    {
        string command = $"\"{Path.GetFullPath(executable)}\" {TrayArgument}";
        if (command.Length > 260)
            throw new IOException("The Windows startup command exceeds 260 characters.");
        return command;
    }

    public static void SetEnabled(bool enabled, string executable)
    {
        if (enabled)
        {
            string command = CommandFor(KeepStartupCopy(executable));
            using var key = Registry.CurrentUser.CreateSubKey(RunKey);
            key.SetValue(EntryName, command, RegistryValueKind.String);
        }
        else
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: true);
            key?.DeleteValue(EntryName, throwOnMissingValue: false);
        }
    }

    public static void UpdateForNewerVersion(string executable)
    {
        string? current = ReadCommand();
        if (current is null || !current.StartsWith('"') || !current.EndsWith(CommandSuffix, StringComparison.OrdinalIgnoreCase))
            return;

        string registered = current[1..^CommandSuffix.Length];
        if (!Path.IsPathFullyQualified(registered) || string.Equals(registered, executable, StringComparison.OrdinalIgnoreCase))
            return;

        var runningVersion = VersionOf(FileVersionInfo.GetVersionInfo(executable));
        if (File.Exists(registered))
        {
            var oldFile = FileVersionInfo.GetVersionInfo(registered);
            if (oldFile.ProductName != "NetSwitch Pro" || VersionOf(oldFile) >= runningVersion)
                return;
        }

        string next = CommandFor(KeepStartupCopy(executable));
        using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: true);
        if (string.Equals(key?.GetValue(EntryName) as string, current, StringComparison.Ordinal))
            key!.SetValue(EntryName, next, RegistryValueKind.String);
    }

    private static Version VersionOf(FileVersionInfo file) =>
        new(file.FileMajorPart, file.FileMinorPart, file.FileBuildPart, file.FilePrivatePart);

    private static string KeepStartupCopy(string executable)
    {
        var version = VersionOf(FileVersionInfo.GetVersionInfo(executable));
        string installs = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Programs", "NetSwitchPro");
        string directory = Path.Combine(installs, version.ToString(3));
        string destination = Path.Combine(directory, "NetSwitchPro.exe");
        if (string.Equals(destination, Path.GetFullPath(executable), StringComparison.OrdinalIgnoreCase))
            return destination;

        Directory.CreateDirectory(directory);
        if (!File.Exists(destination))
        {
            string temporary = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.Copy(executable, temporary);
                try { File.Move(temporary, destination, overwrite: false); }
                catch (IOException) when (File.Exists(destination)) { }
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }
        }

        var saved = FileVersionInfo.GetVersionInfo(destination);
        if (saved.ProductName != "NetSwitch Pro" || VersionOf(saved) != version)
            throw new IOException("The existing NetSwitch Pro startup copy has an unexpected version.");
        return destination;
    }
}
