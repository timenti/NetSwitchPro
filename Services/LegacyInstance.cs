using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace NetworkAdapterSwitcher.Services;

internal static class LegacyInstance
{
    // Published executables that predate the cooperative retirement event.
    private static readonly HashSet<string> KnownBuilds = new(StringComparer.OrdinalIgnoreCase)
    {
        "313c2b990e0910ddff8458ebb867a04ec8f058bc83a5dec892ba93b6b6ebc414", // 2.0.9
        "753768ff586e723759191aeb68eeb399b72dced41d21f0d619ceac2bfae3dfaa", // 2.0.8
        "5433aa691d183fa42f180308493f9e0fd6b6ddc3db31777d7cc562496ae37c24", // 2.0.7
        "321e2ec194adfd8be1f3ca9785ef6de855faef22c97ecf38dc12a30c2576a9e5", // 2.0.6
        "f91f3c7360c752552c5112e765c85ac5369e58700ed3e4c6c77ecb6dba1e723a", // 2.0.5
        "e78a1ef8c068b8433a34ed914bb74b0511c96d4c0207200898eb114987cc0c4b", // 2.0.4
        "8218fafed50bfeb56dbeeab9d4d2b530d46c03d06cedbcc28c1a18567655758a", // 2.0.3
        "c43a31f26a1d9f82f3d30f04d0ff4c9c192dbd2af1642b08938bf45381a571ca", // 2.0.2
        "54bdb063266480615c7d0f501872c1b86c76dc0e8cbe18ee09376e9bca58e307", // 2.0.1
        "412ff00af21d1839fe430a783192249d18ff13c371796e1231ff2c31109b0d55"  // 2.0
    };

    internal static bool RequestExit(string slot, Version incoming)
    {
        var seen = new HashSet<uint>();
        using var current = Process.GetCurrentProcess();
        using var owner = WindowsIdentity.GetCurrent();
        IntPtr window = IntPtr.Zero;
        while ((window = FindWindowEx(IntPtr.Zero, window, null, "NetSwitch Pro")) != IntPtr.Zero)
        {
            GetWindowThreadProcessId(window, out uint pid);
            if (pid == 0 || pid == Environment.ProcessId || !seen.Add(pid))
                continue;
            try
            {
                using var process = Process.GetProcessById((int)pid);
                if (process.SessionId != current.SessionId || !SameOwner(process, owner.User))
                    continue;
                string? executable = process.MainModule?.FileName;
                if (executable is null)
                    continue;
                var metadata = FileVersionInfo.GetVersionInfo(executable);
                if (!StartupRegistration.IsNetSwitch(metadata) || StartupRegistration.VersionOf(metadata) >= incoming)
                    continue;
                using var file = File.OpenRead(executable);
                if (!KnownBuilds.Contains(Convert.ToHexString(SHA256.HashData(file))))
                    continue;

                CloseApplication(process);
                if (RunningInstance.WaitUntilFree(slot, TimeSpan.FromSeconds(2)))
                    return true;
            }
            catch (Exception e) when (e is ArgumentException or InvalidOperationException or Win32Exception
                or IOException or UnauthorizedAccessException)
            {
                // The old process may have exited or Windows may deny access to it.
            }
        }
        return RunningInstance.WaitUntilFree(slot, TimeSpan.Zero);
    }

    private static bool SameOwner(Process process, SecurityIdentifier? owner)
    {
        const uint tokenQuery = 8;
        if (owner is null || !OpenProcessToken(process.SafeHandle, tokenQuery, out var token))
            return false;
        using (token)
        using (var identity = new WindowsIdentity(token.DangerousGetHandle()))
            return identity.User is { } processOwner && owner.Equals(processOwner);
    }

    private static void CloseApplication(Process process)
    {
        if (RmStartSession(out uint session, 0, new StringBuilder(33)) != 0)
            return;
        try
        {
            long started = process.StartTime.ToUniversalTime().ToFileTimeUtc();
            var application = new RegisteredProcess
            {
                Id = (uint)process.Id,
                StartedLow = unchecked((uint)started),
                StartedHigh = (uint)(started >> 32)
            };
            if (RmRegisterResources(session, 0, IntPtr.Zero, 1, [application], 0, IntPtr.Zero) == 0)
                RmShutdown(session, 0, IntPtr.Zero);
        }
        finally { RmEndSession(session); }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RegisteredProcess
    {
        public uint Id, StartedLow, StartedHigh;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string? className, string title);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);

    [DllImport("advapi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool OpenProcessToken(SafeProcessHandle process, uint access, out SafeAccessTokenHandle token);

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern uint RmStartSession(out uint session, uint flags, StringBuilder key);

    [DllImport("rstrtmgr.dll", ExactSpelling = true)]
    private static extern uint RmRegisterResources(uint session, uint files, IntPtr filenames, uint applications,
        [In] RegisteredProcess[] processes, uint services, IntPtr serviceNames);

    [DllImport("rstrtmgr.dll")]
    private static extern uint RmShutdown(uint session, uint flags, IntPtr callback);

    [DllImport("rstrtmgr.dll")]
    private static extern uint RmEndSession(uint session);
}
