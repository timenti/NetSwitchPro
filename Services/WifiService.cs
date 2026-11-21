using System.Collections.Concurrent;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace NetworkAdapterSwitcher.Services;

internal sealed record WifiAdapter(Guid Id, string Description, string[] Profiles, string? ActiveProfile, bool Connected);
internal enum WifiSwitchResult { Connected, NotConfirmed, PreviousRestored, RestoreUnconfirmed }
internal sealed record WifiSwitchOutcome(WifiSwitchResult Result, string? Diagnostic = null);

internal sealed class WifiService : IDisposable
{
    private const uint Acm = 8;
    private const uint ConnectionComplete = 10, ConnectionAttemptFailed = 11, InterfaceRemoval = 14, Disconnected = 21;
    private const int ConnectedState = 1, CurrentConnection = 7;
    private readonly ConcurrentDictionary<Guid, string> observedProfiles = new();
    private readonly Native.NotificationCallback callback;
    private IntPtr handle;
    private Attempt? pending;

    public WifiService() => callback = ObserveNotification;

    public static string? ChooseTarget(string? first, string? second, string? active) =>
        first is null || second is null || first == second ? null : active == first ? second : first;

    public IReadOnlyList<WifiAdapter> GetAdapters(bool readCurrent = true)
    {
        Open();
        Check(Native.WlanEnumInterfaces(handle, IntPtr.Zero, out var data));
        try
        {
            return ReadList<InterfaceInfo>(data).Select(item =>
            {
                string? active = item.State == ConnectedState && readCurrent ? GetCurrent(item.Id) : null;
                return new WifiAdapter(item.Id, item.Description, GetProfiles(item.Id), active, item.State == ConnectedState);
            }).ToArray();
        }
        finally { Native.WlanFreeMemory(data); }
    }

    private string[] GetProfiles(Guid id)
    {
        Check(Native.WlanGetProfileList(handle, ref id, IntPtr.Zero, out var data));
        try { return ReadList<ProfileInfo>(data).Select(info => info.Name).Distinct(StringComparer.Ordinal).ToArray(); }
        finally { Native.WlanFreeMemory(data); }
    }

    private string? GetCurrent(Guid id)
    {
        uint result = Native.WlanQueryInterface(handle, ref id, CurrentConnection, IntPtr.Zero, out var size, out var data, out _);
        if (result != 0)
            return result == 5 && observedProfiles.TryGetValue(id, out var name) ? name : null;
        try
        {
            if (size < 520 || Marshal.ReadInt32(data) != ConnectedState)
                return null;
            return Marshal.PtrToStringUni(IntPtr.Add(data, 8), 256)?.TrimEnd('\0');
        }
        finally { Native.WlanFreeMemory(data); }
    }

    public async Task<WifiSwitchOutcome> SwitchAsync(Guid id, string first, string second, string requested)
    {
        var start = GetAdapters().SingleOrDefault(a => a.Id == id);
        if (start is null || !start.Profiles.Contains(first) || !start.Profiles.Contains(second) ||
            requested != ChooseTarget(first, second, start.ActiveProfile))
            return new(WifiSwitchResult.NotConfirmed);

        return await ConnectSavedAsync(start, requested);
    }

    public Task<WifiSwitchOutcome> ConnectProfileAsync(Guid id, string requested)
    {
        var adapter = GetAdapters().SingleOrDefault(item => item.Id == id);
        return adapter is not null && adapter.Profiles.Contains(requested, StringComparer.Ordinal)
            ? ConnectSavedAsync(adapter, requested)
            : Task.FromResult(new WifiSwitchOutcome(WifiSwitchResult.NotConfirmed));
    }

    private async Task<WifiSwitchOutcome> ConnectSavedAsync(WifiAdapter start, string requested)
    {
        if (start.Connected && start.ActiveProfile == requested)
            return new(WifiSwitchResult.Connected);
        Guid id = start.Id;
        string? previous = start.ActiveProfile;
        var attempt = await ConnectAsync(id, requested);
        if (attempt.Joined)
            return new(WifiSwitchResult.Connected);

        if (!attempt.Accepted || string.IsNullOrEmpty(previous) || previous == requested || !start.Profiles.Contains(previous))
            return new(WifiSwitchResult.NotConfirmed, attempt.Diagnostic);

        var now = GetAdapters().SingleOrDefault(a => a.Id == id);
        if (now?.ActiveProfile == previous)
            return new(WifiSwitchResult.PreviousRestored, attempt.Diagnostic);
        if (now is null || (now.Connected && now.ActiveProfile != requested))
            return new(WifiSwitchResult.RestoreUnconfirmed, attempt.Diagnostic);

        var restored = await ConnectAsync(id, previous);
        return new(restored.Joined ? WifiSwitchResult.PreviousRestored : WifiSwitchResult.RestoreUnconfirmed, attempt.Diagnostic);
    }

    private async Task<(bool Accepted, bool Joined, string? Diagnostic)> ConnectAsync(Guid id, string profile)
    {
        var attempt = new Attempt(id, profile);
        Volatile.Write(ref pending, attempt);
        try
        {
            var settings = new ConnectionParameters { Mode = 0, Profile = profile, BssType = 3 };
            var error = Native.WlanConnect(handle, ref id, ref settings, IntPtr.Zero);
            if (error != 0)
                return (false, false, new Win32Exception((int)error).Message);

            for (int i = 0; i < 36; i++)
            {
                if (attempt.Completed.Task.IsCompletedSuccessfully)
                {
                    uint reason = await attempt.Completed.Task;
                    return reason == 0 ? (true, true, null) : (true, false, Reason(reason));
                }
                if (i % 3 == 0 && GetCurrent(id) == profile)
                    return (true, true, null);
                await Task.Delay(500);
            }
            return (true, false, Reason(Volatile.Read(ref attempt.LastFailure)));
        }
        finally { Volatile.Write(ref pending, null); }
    }

    private void Open()
    {
        if (handle != IntPtr.Zero)
            return;
        Check(Native.WlanOpenHandle(2, IntPtr.Zero, out _, out handle));
        // ACM notifications can confirm a profile switch even when location rules block querying the active profile.
        _ = Native.WlanRegisterNotification(handle, Acm, false, callback, IntPtr.Zero, IntPtr.Zero, out _);
    }

    private void ObserveNotification(ref Notification value, IntPtr context)
    {
        if (value.Source != Acm)
            return;

        if (value.Code is InterfaceRemoval or Disconnected)
        {
            observedProfiles.TryRemove(value.InterfaceId, out _);
            return;
        }

        if (value.Code is not ConnectionComplete and not ConnectionAttemptFailed ||
            value.Data == IntPtr.Zero || value.Size < 564)
            return;

        string profile = Marshal.PtrToStringUni(IntPtr.Add(value.Data, 4), 256)?.TrimEnd('\0') ?? "";
        uint reason = unchecked((uint)Marshal.ReadInt32(value.Data, 560));
        if (value.Code == ConnectionComplete && reason == 0 && profile.Length != 0)
            observedProfiles[value.InterfaceId] = profile;

        var attempt = Volatile.Read(ref pending);
        if (attempt is null || attempt.Id != value.InterfaceId || attempt.Profile != profile)
            return;
        if (value.Code == ConnectionComplete)
            attempt.Completed.TrySetResult(reason);
        else
            Volatile.Write(ref attempt.LastFailure, reason);
    }

    private static string? Reason(uint reason)
    {
        if (reason == 0) return null;
        var buffer = new System.Text.StringBuilder(512);
        return Native.WlanReasonCodeToString(reason, buffer.Capacity, buffer, IntPtr.Zero) == 0
            ? buffer.ToString() : $"WLAN 0x{reason:X}";
    }

    private static IEnumerable<T> ReadList<T>(IntPtr data) where T : struct
    {
        int count = Marshal.ReadInt32(data);
        if (count is < 0 or > 4096)
            throw new InvalidDataException("Invalid Windows WLAN list size.");
        int stride = Marshal.SizeOf<T>();
        for (int i = 0; i < count; i++)
            yield return Marshal.PtrToStructure<T>(IntPtr.Add(data, 8 + i * stride));
    }

    private static void Check(uint code)
    {
        if (code != 0) throw new Win32Exception((int)code);
    }

    public void Dispose()
    {
        if (handle == IntPtr.Zero) return;
        _ = Native.WlanRegisterNotification(handle, 0, false, null, IntPtr.Zero, IntPtr.Zero, out _);
        _ = Native.WlanCloseHandle(handle, IntPtr.Zero);
        handle = IntPtr.Zero;
        GC.KeepAlive(callback);
    }

    private sealed class Attempt(Guid id, string profile)
    {
        public Guid Id { get; } = id;
        public string Profile { get; } = profile;
        public TaskCompletionSource<uint> Completed { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public uint LastFailure;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct InterfaceInfo
    {
        public Guid Id;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string Description;
        public int State;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ProfileInfo
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string Name;
        public uint Flags;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ConnectionParameters
    {
        public int Mode;
        [MarshalAs(UnmanagedType.LPWStr)] public string Profile;
        public IntPtr Ssid;
        public IntPtr DesiredBssids;
        public int BssType;
        public uint Flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Notification
    {
        public uint Source, Code;
        public Guid InterfaceId;
        public uint Size;
        public IntPtr Data;
    }

    private static class Native
    {
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate void NotificationCallback(ref Notification notification, IntPtr context);
        [DllImport("wlanapi.dll")] internal static extern uint WlanOpenHandle(uint version, IntPtr reserved, out uint actual, out IntPtr handle);
        [DllImport("wlanapi.dll")] internal static extern uint WlanCloseHandle(IntPtr handle, IntPtr reserved);
        [DllImport("wlanapi.dll")] internal static extern void WlanFreeMemory(IntPtr data);
        [DllImport("wlanapi.dll")] internal static extern uint WlanEnumInterfaces(IntPtr handle, IntPtr reserved, out IntPtr data);
        [DllImport("wlanapi.dll")] internal static extern uint WlanGetProfileList(IntPtr handle, ref Guid id, IntPtr reserved, out IntPtr data);
        [DllImport("wlanapi.dll")] internal static extern uint WlanQueryInterface(IntPtr handle, ref Guid id, int opcode, IntPtr reserved, out uint size, out IntPtr data, out int value);
        [DllImport("wlanapi.dll")] internal static extern uint WlanConnect(IntPtr handle, ref Guid id, ref ConnectionParameters settings, IntPtr reserved);
        [DllImport("wlanapi.dll")] internal static extern uint WlanRegisterNotification(IntPtr handle, uint source, [MarshalAs(UnmanagedType.Bool)] bool ignoreDuplicates, NotificationCallback? callback, IntPtr context, IntPtr reserved, out uint previous);
        [DllImport("wlanapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)] internal static extern uint WlanReasonCodeToString(uint reason, int size, System.Text.StringBuilder text, IntPtr reserved);
    }
}
