using System.ComponentModel;
using NetworkAdapterSwitcher.Models;

namespace NetworkAdapterSwitcher.Services;

internal enum AdapterSwitchResult
{
    Changed, TargetNotReady, WifiProfileRequired, WifiProfileUnavailable, WifiNotConnected, FailedRestored, RestoreFailed
}
internal sealed record AdapterSwitchOutcome(AdapterSwitchResult Result, string? Diagnostic = null);

internal sealed class AdapterSwitch(
    IAdapterService adapters, Func<TimeSpan, Task>? delay = null, IWifiProfileService? wifi = null)
{
    private readonly Func<TimeSpan, Task> wait = delay ?? Task.Delay;

    public static NetworkAdapterInfo ChooseTarget(NetworkAdapterInfo first, NetworkAdapterInfo second)
    {
        if (first.Enabled != second.Enabled)
            return first.Enabled ? second : first;
        if (first.Preferred != second.Preferred)
            return first.Preferred ? second : first;
        if (first.Connected != second.Connected)
            return first.Connected ? second : first;
        if (first.Enabled && first.IsWifi != second.IsWifi)
            return first.IsWifi ? first : second;
        return first.Enabled ? second : first;
    }

    public async Task<AdapterSwitchOutcome> ApplyAsync(Guid first, Guid second, Guid? requestedTarget = null, string? wifiProfile = null)
    {
        if (first == second)
            throw new ArgumentException("Choose two different adapters.");
        if (requestedTarget.HasValue && requestedTarget != first && requestedTarget != second)
            throw new ArgumentException("The target must be part of the selected adapter pair.");

        var initial = adapters.GetAdapters().ToDictionary(a => a.Id);
        if (!initial.TryGetValue(first, out var a) || !initial.TryGetValue(second, out var b))
            return new(AdapterSwitchResult.TargetNotReady);
        var target = requestedTarget is null ? ChooseTarget(a, b) : initial[requestedTarget.Value];
        var source = target.Id == a.Id ? b : a;
        bool targetTouched = false;
        bool sourceTouched = false;
        bool profileChanged = false;
        string? previousProfile = null;

        try
        {
            if (target.IsWifi && target.Enabled && wifiProfile is not null)
                previousProfile = FindWifi(target.Id)?.ActiveProfile;

            if (!target.Enabled)
            {
                targetTouched = true;
                adapters.SetEnabled(target.Id, true);
            }
            if (!await ObserveAsync(target.Id, adapter => adapter.Enabled, 20))
                return await RestoreOutcome(AdapterSwitchResult.TargetNotReady);

            if (target.IsWifi && wifiProfile is not null)
            {
                profileChanged = previousProfile is not null && previousProfile != wifiProfile;
                var outcome = await ConnectWifiAsync(target.Id, wifiProfile);
                if (outcome.Result != WifiSwitchResult.Connected &&
                    outcome.RetryWithoutEthernet && source.Enabled && source.Connected && source.Type == "Ethernet")
                {
                    // Some Windows configurations suppress WLAN while wired connectivity is present.
                    await DisableSourceAsync();
                    outcome = await ConnectWifiAsync(target.Id, wifiProfile);
                }

                if (outcome.Result != WifiSwitchResult.Connected)
                {
                    var failure = outcome.Result == WifiSwitchResult.ProfileUnavailable
                        ? AdapterSwitchResult.WifiProfileUnavailable
                        : AdapterSwitchResult.WifiNotConnected;
                    return await RestoreOutcome(failure, outcome.Diagnostic);
                }
            }

            bool Ready(NetworkAdapterInfo current) => current.Enabled &&
                (target.IsWifi
                    ? WifiReady(current, wifiProfile)
                    : !source.Connected || current.Connected);

            if (!await ObserveAsync(target.Id, Ready, 20))
                return await RestoreOutcome(target.IsWifi
                    ? wifiProfile is null ? AdapterSwitchResult.WifiProfileRequired : AdapterSwitchResult.WifiNotConnected
                    : AdapterSwitchResult.TargetNotReady);

            await DisableSourceAsync();

            if (!await ObserveAsync(target.Id, Ready, 5))
                throw new InvalidOperationException("Windows did not confirm the new adapter.");

            return new(AdapterSwitchResult.Changed);
        }
        catch (Exception ex)
        {
            return await RestoreOutcome(AdapterSwitchResult.FailedRestored, ex.Message);
        }

        async Task DisableSourceAsync()
        {
            if (!source.Enabled || sourceTouched)
                return;
            sourceTouched = true;
            adapters.SetEnabled(source.Id, false);
            if (!await ObserveAsync(source.Id, adapter => !adapter.Enabled, 10))
                throw new InvalidOperationException("Windows did not confirm that the old adapter was disabled.");
        }

        async Task<AdapterSwitchOutcome> RestoreOutcome(AdapterSwitchResult failure, string? diagnostic = null)
        {
            bool restored = await RestoreAsync(source, target, sourceTouched, targetTouched, profileChanged ? previousProfile : null);
            return new(restored ? failure : AdapterSwitchResult.RestoreFailed, diagnostic);
        }
    }

    private WifiAdapter? FindWifi(Guid id)
    {
        try { return wifi?.FindAdapter(id); }
        catch (Win32Exception) { return null; }
    }

    private bool WifiReady(NetworkAdapterInfo adapter, string? profile)
    {
        if (profile is null && adapter.Connected)
            return true;
        var active = FindWifi(adapter.Id);
        return active?.Connected == true && (profile is null || active.ActiveProfile == profile);
    }

    private async Task<WifiSwitchOutcome> ConnectWifiAsync(Guid id, string profile)
    {
        if (wifi is null)
            return new(WifiSwitchResult.ProfileUnavailable);
        for (int attempt = 0; attempt < 10; attempt++)
        {
            var available = FindWifi(id);
            if (available?.Profiles.Contains(profile, StringComparer.Ordinal) == true)
                return await wifi.ConnectProfileAsync(id, profile);
            if (available?.Profiles.Length > 0)
                break;
            await wait(TimeSpan.FromMilliseconds(350));
        }
        return new(WifiSwitchResult.ProfileUnavailable);
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

    private async Task<bool> RestoreAsync(NetworkAdapterInfo source, NetworkAdapterInfo target,
        bool sourceTouched, bool targetTouched, string? previousProfile)
    {
        try
        {
            if (!sourceTouched && !targetTouched && previousProfile is null)
                return true;

            var current = adapters.GetAdapters().ToDictionary(a => a.Id);
            if (sourceTouched && source.Enabled &&
                (!current.TryGetValue(source.Id, out var old) || !old.Enabled))
                adapters.SetEnabled(source.Id, true);

            // Leave any replacement path enabled if recovering the old one cannot be confirmed.
            if (source.Enabled && (sourceTouched || source.Connected) &&
                !await ObserveAsync(source.Id, a => a.Enabled && (!source.Connected || a.Connected), 20))
                return false;

            if (targetTouched && !target.Enabled && adapters.GetAdapters().Any(a => a.Id == target.Id && a.Enabled))
            {
                adapters.SetEnabled(target.Id, false);
                return await ObserveAsync(target.Id, a => !a.Enabled, 10);
            }

            return previousProfile is null ||
                (await ConnectWifiAsync(target.Id, previousProfile)).Result == WifiSwitchResult.Connected;
        }
        catch
        {
            return false;
        }
    }
}
