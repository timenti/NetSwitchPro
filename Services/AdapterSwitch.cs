using NetworkAdapterSwitcher.Models;

namespace NetworkAdapterSwitcher.Services;

internal enum AdapterSwitchResult { Changed, TargetNotReady, FailedRestored, RestoreFailed }
internal sealed record AdapterSwitchOutcome(AdapterSwitchResult Result, string? Diagnostic = null);

internal sealed class AdapterSwitch(IAdapterService adapters, Func<TimeSpan, Task>? delay = null)
{
    private readonly Func<TimeSpan, Task> wait = delay ?? Task.Delay;

    public static NetworkAdapterInfo ChooseTarget(NetworkAdapterInfo first, NetworkAdapterInfo second)
    {
        if (first.Enabled != second.Enabled)
            return first.Enabled ? second : first;
        if (first.Connected != second.Connected)
            return first.Connected ? second : first;
        return first.Enabled ? second : first;
    }

    public async Task<AdapterSwitchOutcome> ApplyAsync(Guid first, Guid second)
    {
        if (first == second)
            throw new ArgumentException("Choose two different adapters.");

        var initial = adapters.GetAdapters().ToDictionary(a => a.Id);
        if (!initial.TryGetValue(first, out var a) || !initial.TryGetValue(second, out var b))
            return new(AdapterSwitchResult.TargetNotReady);
        var target = ChooseTarget(a, b);
        var source = target.Id == a.Id ? b : a;
        bool targetTouched = false;
        bool sourceTouched = false;

        try
        {
            if (!target.Enabled)
            {
                targetTouched = true;
                adapters.SetEnabled(target.Id, true);
            }

            // Keep the working link until its replacement has an active link.
            if (!await ObserveAsync(target.Id,
                    adapter => adapter.Enabled && (!source.Connected || adapter.Connected), 20))
            {
                return await RestoreAsync(source, target, sourceTouched, targetTouched)
                    ? new(AdapterSwitchResult.TargetNotReady)
                    : new(AdapterSwitchResult.RestoreFailed);
            }

            if (source.Enabled)
            {
                sourceTouched = true;
                adapters.SetEnabled(source.Id, false);
                if (!await ObserveAsync(source.Id, adapter => !adapter.Enabled, 10))
                    throw new InvalidOperationException("Windows did not confirm that the old adapter was disabled.");
            }

            if (!await ObserveAsync(target.Id, adapter => adapter.Enabled && (!source.Connected || adapter.Connected), 5))
                throw new InvalidOperationException("Windows did not confirm the new adapter.");

            return new(AdapterSwitchResult.Changed);
        }
        catch (Exception ex)
        {
            return await RestoreAsync(source, target, sourceTouched, targetTouched)
                ? new(AdapterSwitchResult.FailedRestored, ex.Message)
                : new(AdapterSwitchResult.RestoreFailed, ex.Message);
        }
    }

    private async Task<bool> ObserveAsync(Guid id, Func<NetworkAdapterInfo, bool> matches, int attempts)
    {
        for (var attempt = 0; attempt < attempts; attempt++)
        {
            if (adapters.GetAdapters().FirstOrDefault(a => a.Id == id) is { } current && matches(current))
                return true;
            await wait(TimeSpan.FromMilliseconds(400));
        }
        return false;
    }

    private async Task<bool> RestoreAsync(NetworkAdapterInfo source, NetworkAdapterInfo target, bool sourceTouched, bool targetTouched)
    {
        try
        {
            if (!sourceTouched && !targetTouched)
                return true;

            var current = adapters.GetAdapters().ToDictionary(a => a.Id);
            if (sourceTouched && source.Enabled &&
                (!current.TryGetValue(source.Id, out var old) || !old.Enabled))
                adapters.SetEnabled(source.Id, true);

            if (source.Enabled && (sourceTouched || source.Connected) &&
                !await ObserveAsync(source.Id, a => a.Enabled && (!source.Connected || a.Connected), 20))
                return false;

            if (targetTouched && !target.Enabled &&
                adapters.GetAdapters().Any(a => a.Id == target.Id && a.Enabled))
            {
                adapters.SetEnabled(target.Id, false);
                return await ObserveAsync(target.Id, a => !a.Enabled, 10);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }
}
