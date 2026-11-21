using System.IO.MemoryMappedFiles;

namespace NetworkAdapterSwitcher.Services;

internal sealed class RunningInstance : IDisposable
{
    private readonly MemoryMappedFile metadata;

    public RunningInstance(string name, Version version)
    {
        Show = new EventWaitHandle(false, EventResetMode.AutoReset, name + ".show");
        Retire = new EventWaitHandle(false, EventResetMode.AutoReset, name + ".retire");
        metadata = MemoryMappedFile.CreateNew(name + ".version", sizeof(int) * 4);
        using var view = metadata.CreateViewAccessor();
        view.Write(0, version.Major);
        view.Write(4, version.Minor);
        view.Write(8, version.Build);
        view.Write(12, version.Revision);
    }

    public EventWaitHandle Show { get; }
    public EventWaitHandle Retire { get; }

    public static Version? ReadVersion(string name)
    {
        try
        {
            using var metadata = MemoryMappedFile.OpenExisting(name + ".version", MemoryMappedFileRights.Read);
            using var view = metadata.CreateViewAccessor(0, 0, MemoryMappedFileAccess.Read);
            var version = new Version(view.ReadInt32(0), view.ReadInt32(4), view.ReadInt32(8), view.ReadInt32(12));
            return version.Major > 0 ? version : null;
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return null;
        }
    }

    public static bool WaitUntilFree(string name, TimeSpan timeout)
    {
        try
        {
            if (!Mutex.TryOpenExisting(name, out var mutex))
                return true;
            using (mutex)
            {
                if (!TryEnter(mutex, timeout))
                    return false;
                mutex.ReleaseMutex();
                return true;
            }
        }
        catch (UnauthorizedAccessException) { return false; }
    }

    public static bool TryEnter(Mutex mutex, TimeSpan timeout)
    {
        try { return mutex.WaitOne(timeout); }
        catch (AbandonedMutexException) { return true; }
    }

    public static bool Signal(string name, string signal)
    {
        try
        {
            if (!EventWaitHandle.TryOpenExisting(name + signal, out var handle))
                return false;
            using (handle)
                return handle.Set();
        }
        catch (UnauthorizedAccessException) { return false; }
    }

    public void Dispose()
    {
        metadata.Dispose();
        Show.Dispose();
        Retire.Dispose();
    }
}
