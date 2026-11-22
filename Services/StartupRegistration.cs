using System.Diagnostics;
using Microsoft.Win32;

namespace NetworkAdapterSwitcher.Services;

internal static class StartupRegistration
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    internal const string EntryName = "NetSwitchPro";
    internal const string TrayArgument = "--start-in-tray";

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
            string command = CommandFor(executable);
            using var key = Registry.CurrentUser.CreateSubKey(RunKey);
            key.SetValue(EntryName, command, RegistryValueKind.String);
        }
        else
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: true);
            key?.DeleteValue(EntryName, throwOnMissingValue: false);
        }
    }

    internal static Version VersionOf(FileVersionInfo file) =>
        new(file.FileMajorPart, file.FileMinorPart, file.FileBuildPart, file.FilePrivatePart);

    internal static bool IsNetSwitch(FileVersionInfo file) =>
        file.ProductName == "NetSwitch Pro" && file.OriginalFilename == "NetworkAdapterSwitcher.dll";
}
