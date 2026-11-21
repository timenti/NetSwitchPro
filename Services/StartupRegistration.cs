using System.Diagnostics;
using Microsoft.Win32;

namespace NetworkAdapterSwitcher.Services;

internal static class StartupRegistration
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    internal const string EntryName = "NetSwitchPro";
    internal const string TrayArgument = "--start-in-tray";
    internal static string ProgramDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "NetSwitchPro");
    internal static string CurrentPath => Path.Combine(ProgramDirectory, "Current", "NetSwitchPro.exe");
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
            string destination = UpdatePublisher.IsTrusted(executable) ? InstallCurrentCopy(executable) : executable;
            string command = CommandFor(destination);
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
        bool managedCopyExists = File.Exists(CurrentPath);
        if (!managedCopyExists && current is null)
            return;
        if (string.Equals(executable, CurrentPath, StringComparison.OrdinalIgnoreCase) &&
            (current is null || current == CommandFor(CurrentPath)))
            return;
        if (!UpdatePublisher.IsTrusted(executable))
            return;
        if (managedCopyExists)
            InstallCurrentCopy(executable);
        if (current is null || !current.StartsWith('"') || !current.EndsWith(CommandSuffix, StringComparison.OrdinalIgnoreCase))
            return;

        string registered = current[1..^CommandSuffix.Length];
        if (!Path.IsPathFullyQualified(registered) || string.Equals(registered, executable, StringComparison.OrdinalIgnoreCase))
            return;

        var runningVersion = VersionOf(FileVersionInfo.GetVersionInfo(executable));
        if (File.Exists(registered))
        {
            var oldFile = FileVersionInfo.GetVersionInfo(registered);
            if (!IsNetSwitch(oldFile) || VersionOf(oldFile) > runningVersion)
                return;
        }

        string next = CommandFor(InstallCurrentCopy(executable));
        using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: true);
        if (string.Equals(key?.GetValue(EntryName) as string, current, StringComparison.Ordinal))
            key!.SetValue(EntryName, next, RegistryValueKind.String);
    }

    internal static Version VersionOf(FileVersionInfo file) =>
        new(file.FileMajorPart, file.FileMinorPart, file.FileBuildPart, file.FilePrivatePart);

    internal static bool IsNetSwitch(FileVersionInfo file) =>
        file.ProductName == "NetSwitch Pro" && file.OriginalFilename == "NetworkAdapterSwitcher.dll";

    internal static string InstallCurrentCopy(string executable)
    {
        string source = Path.GetFullPath(executable);
        string destination = CurrentPath;
        if (string.Equals(destination, source, StringComparison.OrdinalIgnoreCase))
            return destination;

        var incoming = FileVersionInfo.GetVersionInfo(source);
        if (!IsNetSwitch(incoming))
            throw new IOException("This file is not a NetSwitch Pro build.");
        if (File.Exists(destination))
        {
            var existing = FileVersionInfo.GetVersionInfo(destination);
            if (!IsNetSwitch(existing))
                throw new IOException("The NetSwitch Pro installation path contains another application.");
            if (VersionOf(existing) >= VersionOf(incoming))
                return destination;
        }

        var directory = Path.GetDirectoryName(destination)!;
        Directory.CreateDirectory(directory);
        string temporary = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            File.Copy(source, temporary);
            File.Move(temporary, destination, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary))
                File.Delete(temporary);
        }
        return destination;
    }
}
