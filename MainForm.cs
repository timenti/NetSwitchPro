using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using NetworkAdapterSwitcher.Models;
using NetworkAdapterSwitcher.Services;

namespace NetworkAdapterSwitcher;

internal sealed class MainForm : Form
{
    private readonly AppSettings settings = AppSettings.Load();
    private readonly WifiService wifi = new();
    private readonly IAdapterService adapters = new WindowsAdapterService();
    private readonly bool administrator = IsAdministrator();
    private readonly List<string> journal = [];
    private readonly Label category = TextLabel(11, "muted", bold: true);
    private readonly Label activeLabel = TextLabel(10, "muted");
    private readonly Label current = TextLabel(23, bold: true);
    private readonly Label pairTitle = TextLabel(13, bold: true);
    private readonly Label connectionDetail = TextLabel(10, "muted");
    private readonly Label hint = TextLabel(10, "muted");
    private readonly Label status = TextLabel(10, "muted");
    private readonly DesktopButton refresh = MakeButton(glyph: "");
    private readonly DesktopButton history = MakeButton(glyph: "");
    private readonly DesktopButton preferences = MakeButton(glyph: "");
    private readonly DesktopButton elevate = MakeButton(ButtonAppearance.Outline, glyph: "\uE72E");
    private readonly DesktopButton action = MakeButton(ButtonAppearance.Primary, "");
    private const string AutomaticWifiProfile = "auto";
    private const string LoadWifiNetworks = "load";
    private const string WifiProfilePrefix = "profile:";
    private readonly DesktopToggle showVirtual = new();
    private readonly DesktopToggle showBluetooth = new();
    private readonly PairCard first = new();
    private readonly PairCard second = new();
    private readonly TableLayoutPanel connection = Surface(2, 2);
    private readonly TableLayoutPanel filters = Grid(1, 2);
    private readonly ToolTip tooltips = new();
    private readonly Icon captionIcon = DesktopDrawing.LoadCompactAppIcon(SystemInformation.SmallIconSize);
    private readonly NotifyIcon notification = new() { Text = "NetSwitch Pro" };
    private readonly TrayPanel tray = new();
    private TrayIconFeedback? iconFeedback;
    private bool allowExit;
    private bool upgradeExitRequested;
    private Form? activePreferences;
    private string? trayFeedback;
    private IReadOnlyList<WifiAdapter> wireless = [];
    private IReadOnlyList<NetworkAdapterInfo> wired = [];
    private bool busy;
    private bool hasShown;
    private bool refreshing;
    private long lastRead;
    private bool binding;
    private string? wifiReadError, adapterReadError;
    private bool lightTheme;

    public MainForm()
    {
        lightTheme = settings.IsLight;
        SuspendLayout();
        Text = "NetSwitch Pro";
        AutoScaleDimensions = new SizeF(96, 96);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 10);
        ClientSize = new Size(656, 470);
        MinimumSize = new Size(607, 504);
        StartPosition = FormStartPosition.CenterScreen;
        Icon = captionIcon;
        BuildLayout();
        HandleCreated += (_, _) => SetTitleBar(this);
        UpdateLabels();
        Render();
        PaintTheme(this);
        ResumeLayout(true);

        refresh.Click += async (_, _) => await RefreshAsync();
        history.Click += (_, _) => ShowJournal();
        preferences.Click += (_, _) => ShowPreferences();
        elevate.Click += (_, _) => Elevate();
        action.Click += async (_, _) => await RunConnectionAsync(null, fromTray: false);
        first.Picker.SelectionChangeCommitted += (_, _) => SavePair(first);
        second.Picker.SelectionChangeCommitted += (_, _) => SavePair(second);
        showVirtual.CheckedChanged += (_, _) => ChangeFilters();
        showBluetooth.CheckedChanged += (_, _) => ChangeFilters();
        Shown += async (_, _) =>
        {
            InitializeTray();
            await RefreshAsync();
        };
        Activated += async (_, _) =>
        {
            UpdateSystemTheme();
            if (hasShown && Environment.TickCount64 - lastRead >= 6000)
                await RefreshAsync(background: true);
        };
        notification.MouseClick += async (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
                await RunConnectionAsync(null, fromTray: true);
            else if (e.Button == MouseButtons.Right)
                await OpenTrayAsync(Cursor.Position);
        };
        tray.ProfileRequested += async target => await RunConnectionAsync(target, fromTray: true);
        tray.OpenRequested += RestoreMainWindow;
        tray.ExitRequested += ExitApplication;

        FormClosing += (_, e) =>
        {
            if (e.CloseReason == CloseReason.UserClosing && !allowExit && notification.Visible)
            {
                e.Cancel = true;
                Hide();
            }
        };

    }

    internal void InitializeTray()
    {
        if (hasShown)
            return;
        hasShown = true;
        iconFeedback = new TrayIconFeedback(notification);
        notification.Visible = true;
    }

    private void BuildLayout()
    {
        var root = Grid(1, 7);
        root.Padding = new Padding(24, 19, 24, 16);
        foreach (int height in new[] { 61, 30, 94, 41, 134 })
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 61));
        Controls.Add(root);

        var header = Grid(5, 1);
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (int i = 0; i < 3; i++)
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 43));
        header.Margin = new Padding(0, 0, 0, 12);
        var brand = new PictureBox { Dock = DockStyle.Fill, Margin = new Padding(0, 1, 13, 1), SizeMode = PictureBoxSizeMode.Zoom };
        using (var stream = typeof(MainForm).Assembly.GetManifestResourceStream("NetworkAdapterSwitcher.BrandIcon")!)
        using (var asset = new Icon(stream, 64, 64))
            brand.Image = asset.ToBitmap();
        brand.Disposed += (_, _) => brand.Image?.Dispose();
        var title = TextLabel(24, bold: true);
        title.Text = "NetSwitch Pro";
        header.Controls.Add(brand, 0, 0);
        header.Controls.Add(title, 1, 0);
        header.Controls.Add(refresh, 2, 0);
        header.Controls.Add(history, 3, 0);
        header.Controls.Add(preferences, 4, 0);
        for (int i = 2; i <= 4; i++)
            header.GetControlFromPosition(i, 0)!.Margin = new Padding(5, 6, 0, 6);
        root.Controls.Add(header, 0, 0);

        category.Margin = pairTitle.Margin = new Padding(10, 0, 0, 0);
        root.Controls.Add(category, 0, 1);

        var connectionRow = Grid(2, 1);
        connectionRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        connectionRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 156));
        connectionRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        connectionRow.Margin = new Padding(0, 0, 0, 6);

        connection.Padding = new Padding(16, 9, 16, 10);
        connection.Margin = new Padding(0, 0, 12, 0);
        connection.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        connection.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        connection.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
        connection.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        activeLabel.Margin = new Padding(4, 0, 0, 0);
        connection.Controls.Add(activeLabel, 0, 0);
        connection.SetColumnSpan(activeLabel, 2);
        connection.Controls.Add(current, 0, 1);
        connectionDetail.AutoSize = true;
        connectionDetail.Margin = new Padding(10, 0, 0, 0);
        connectionDetail.TextAlign = ContentAlignment.MiddleRight;
        connection.Controls.Add(connectionDetail, 1, 1);
        connectionRow.Controls.Add(connection, 0, 0);

        filters.Margin = new Padding(0, 5, 0, 5);
        filters.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        filters.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        showVirtual.Dock = showBluetooth.Dock = DockStyle.Fill;
        showVirtual.Margin = new Padding(0, 0, 0, 3);
        showBluetooth.Margin = new Padding(0, 3, 0, 0);
        filters.Controls.Add(showVirtual, 0, 0);
        filters.Controls.Add(showBluetooth, 0, 1);
        connectionRow.Controls.Add(filters, 1, 0);
        root.Controls.Add(connectionRow, 0, 2);
        root.Controls.Add(pairTitle, 0, 3);

        var pair = Surface(1, 3);
        pair.Padding = new Padding(15, 4, 15, 4);
        pair.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        pair.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));
        pair.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        pair.Controls.Add(first.Container, 0, 0);
        pair.Controls.Add(new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty, Tag = "line" }, 0, 1);
        pair.Controls.Add(second.Container, 0, 2);
        root.Controls.Add(pair, 0, 4);

        hint.Margin = status.Margin = new Padding(0, 0, 15, 0);
        var footer = Grid(2, 1);
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 57));
        footer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var feedback = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 15, 0) };
        feedback.Controls.Add(hint);
        feedback.Controls.Add(status);
        footer.Controls.Add(feedback, 0, 0);
        var actions = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 6, 0, 6) };
        actions.Controls.Add(action);
        actions.Controls.Add(elevate);
        footer.Controls.Add(actions, 1, 0);
        root.Controls.Add(footer, 0, 6);
    }

    private string W(Word word) => UiWords.Get(settings.Language, word);
    private string Say(string ru, string en) => settings.Language == "Russian" ? ru : en;

    private void UpdateLabels()
    {
        category.Text = W(Word.NetworksAndAdapters);
        activeLabel.Text = W(Word.TrayCurrentConnection);
        pairTitle.Text = W(Word.PairNetworks);
        first.Caption.Text = W(Word.NetworkOne);
        second.Caption.Text = W(Word.NetworkTwo);
        refresh.AccessibleName = W(Word.Refresh);
        history.AccessibleName = W(Word.Log);
        preferences.AccessibleName = W(Word.Settings);
        foreach (var button in new[] { refresh, history, preferences })
            tooltips.SetToolTip(button, button.AccessibleName);
        showVirtual.Text = W(Word.Virtual);
        showBluetooth.Text = W(Word.Bluetooth);
        elevate.Text = Say("Запросить права", "Run as administrator");
        tooltips.SetToolTip(elevate, W(Word.Elevate));
        action.Text = W(Word.Switch);
    }

    private async Task RefreshAsync(bool background = false)
    {
        if (busy || refreshing)
            return;
        refreshing = true;
        if (background)
            action.Enabled = false;
        else
            SetBusy(true);
        UpdateTray();
        try
        {
            await ReadConnectionsAsync();
        }
        finally
        {
            refreshing = false;
            lastRead = Environment.TickCount64;
            if (background)
                Render();
            else
                SetBusy(false);
            CloseForUpgradeWhenIdle();
        }
    }

    private async Task ReadAsync(bool readWifi)
    {
        string? error = null;
        try
        {
            if (readWifi)
            {
                wireless = await Task.Run(() => wifi.GetAdapters());
                bool changed = false;
                foreach (var adapter in wireless.Where(item => item.Profiles.Length > 0))
                {
                    if (settings.KnownWifiProfiles.TryGetValue(adapter.Id, out var known) &&
                        known.SequenceEqual(adapter.Profiles, StringComparer.Ordinal))
                        continue;
                    settings.KnownWifiProfiles[adapter.Id] = adapter.Profiles;
                    changed = true;
                }
                if (changed)
                    SaveSettings();
            }
            else
                wired = await Task.Run(adapters.GetAdapters);
        }
        catch (Exception ex)
        {
            if (readWifi)
                wireless = [];
            else
                wired = [];
            error = Say("Не удалось прочитать данные Windows. См. журнал.", "Windows data is unavailable. See activity.");
            WriteJournal(ex.Message, showStatus: false);
        }
        if (readWifi)
            wifiReadError = error;
        else
            adapterReadError = error;
    }

    private ConnectionPair Pair => new(settings.NetworkOne, settings.NetworkTwo);

    private NetworkAdapterInfo[] ConnectionAdapters()
    {
        var result = wired.ToDictionary(adapter => adapter.Id);
        foreach (var adapter in wireless)
        {
            if (result.TryGetValue(adapter.Id, out var existing))
            {
                if (adapter.Connected && !existing.Connected)
                    result[adapter.Id] = existing with { Connected = true, Enabled = true };
                continue;
            }
            result[adapter.Id] = new(adapter.Id, adapter.Description, adapter.Description,
                "Wi-Fi", true, adapter.Connected, false, false);
        }
        return result.Values.OrderBy(adapter => adapter.Name, StringComparer.CurrentCultureIgnoreCase).ToArray();
    }

    private bool CanDisplayAdapter(NetworkAdapterInfo adapter) =>
        (settings.ShowVirtual || !adapter.IsVirtual) && (settings.ShowBluetooth || !adapter.IsBluetooth);

    private async Task ReadConnectionsAsync()
    {
        await Task.WhenAll(ReadAsync(true), ReadAsync(false));
        if (settings.UnifiedPairSaved)
            return;

        var inventory = ConnectionAdapters();
        NetworkEndpoint? LegacyAdapter(Guid? id, string? name)
        {
            var resolved = id ?? inventory.FirstOrDefault(adapter => adapter.Name.Equals(name, StringComparison.OrdinalIgnoreCase))?.Id;
            return resolved is { } value ? new(value, settings.AdapterWifiProfiles.GetValueOrDefault(value)) : null;
        }

        var wifiId = settings.WifiAdapterId
            ?? wireless.FirstOrDefault(adapter => adapter.Connected)?.Id
            ?? wireless.FirstOrDefault()?.Id;
        bool hadWifiSelection = settings.WifiA is not null || settings.WifiB is not null;
        var oldWifi = new ConnectionPair(
            wifiId is { } firstWifi && settings.WifiA is { } a ? new(firstWifi, a) : null,
            wifiId is { } secondWifi && settings.WifiB is { } b ? new(secondWifi, b) : null);
        var oldAdapters = new ConnectionPair(LegacyAdapter(settings.AdapterAId, settings.AdapterA),
            LegacyAdapter(settings.AdapterBId, settings.AdapterB));
        if (!settings.ShowAdapters && hadWifiSelection && wifiId is null)
            return;

        var preferred = settings.ShowAdapters ? oldAdapters : oldWifi;
        var fallback = settings.ShowAdapters ? oldWifi : oldAdapters;
        var migrated = preferred.First is not null || preferred.Second is not null ? preferred : fallback;
        if (migrated.First is null && migrated.Second is null)
            return;

        settings.NetworkOne = migrated.First;
        settings.NetworkTwo = migrated.Second;
        settings.UnifiedPairSaved = true;
        SaveSettings();
    }

    private void Render()
    {
        binding = true;
        SuspendLayout();
        try
        {
            var inventory = ConnectionAdapters();
            var pair = Pair;
            var active = pair.Current(inventory, wireless);
            var target = pair.ResolveTarget(inventory, wireless);
            var currentAdapter = FindAdapter(inventory, active);
            bool needsAdmin = pair.UsesDifferentAdapters && !administrator;
            current.Text = active is null ? W(Word.Disconnect) : NameFor(active, inventory);
            tooltips.SetToolTip(current, current.Text);
            connectionDetail.Text = currentAdapter?.IPv4 is { } ipv4 ? "IP  " + ipv4 : "";
            elevate.Visible = needsAdmin;
            action.Visible = !needsAdmin;
            first.Picker.Enabled = second.Picker.Enabled = !busy && !refreshing;
            status.Text = (journal.Count > 0 ? journal[0] : null) ??
                (inventory.Length == 0 ? adapterReadError ?? wifiReadError : null) ?? "";
            showVirtual.Checked = settings.ShowVirtual;
            showBluetooth.Checked = settings.ShowBluetooth;

            var choices = inventory.Where(adapter => CanDisplayAdapter(adapter) || pair.ContainsAdapter(adapter.Id))
                .Select(adapter => new DesktopChoice(adapter.Id.ToString(), adapter.Name, AdapterDetail(adapter),
                    active?.AdapterId == adapter.Id, OpensChoices: adapter.IsWifi)).ToArray();
            BindEndpoint(first, pair.First, choices, inventory, active);
            BindEndpoint(second, pair.Second, choices, inventory, active);

            string? problem = target is null ? PairProblem(pair, inventory) : AvailabilityProblem(target, pair, inventory);
            hint.Text = problem ?? (needsAdmin ? W(Word.ConnectionAdminRequired) : "");
            SetAction(target is null ? null : NameFor(target, inventory), W(Word.SwitchPair));
            action.Enabled = !busy && !refreshing && !needsAdmin && target is not null && problem is null;
            if (busy)
                action.Text = W(Word.Busy);
            ShowFeedback();
            action.Invalidate();
        }
        finally
        {
            ResumeLayout(true);
            binding = false;
        }
        UpdateTray();
    }

    private void BindEndpoint(PairCard card, NetworkEndpoint? endpoint, DesktopChoice[] choices,
        NetworkAdapterInfo[] inventory, NetworkEndpoint? active)
    {
        var adapter = FindAdapter(inventory, endpoint);
        var title = endpoint?.WifiProfile is { } profile && adapter is not null ? adapter.Name + " · " + profile : null;
        card.Picker.Placeholder = endpoint is null ? W(Word.ChooseConnection) : endpoint.WifiProfile ?? W(Word.ConnectionUnavailable);
        card.Picker.Bind(choices, adapter?.Id.ToString(), title);
        card.Detail.Text = endpoint switch
        {
            null => "",
            _ when adapter is null => W(Word.ConnectionUnavailable),
            { WifiProfile: not null } => (ConnectionPair.Matches(active, endpoint) ? W(Word.Connected) : W(Word.Saved)) + "   ·   " + adapter.Type,
            _ => AdapterDetail(adapter)
        };
        tooltips.SetToolTip(card.Picker, title ?? adapter?.Name ?? card.Picker.Placeholder);
        tooltips.SetToolTip(card.Detail, adapter?.Description ?? "");
    }

    private string PairProblem(ConnectionPair pair, NetworkAdapterInfo[] inventory)
    {
        if (!pair.HasBoth)
            return W(Word.PickConnections);
        if (!pair.UsesDifferentAdapters && FindAdapter(inventory, pair.First)?.IsWifi == true)
            return W(Word.DifferentWifiProfiles);
        return pair.First == pair.Second ? W(Word.PickConnections) : W(Word.ConnectionUnavailable);
    }

    private string? AvailabilityProblem(NetworkEndpoint endpoint, ConnectionPair pair, NetworkAdapterInfo[] inventory)
    {
        var adapter = FindAdapter(inventory, endpoint);
        if (adapter is null || (pair.OtherAdapterFor(endpoint) is { } source && !inventory.Any(item => item.Id == source)))
            return W(Word.ConnectionUnavailable);
        if (endpoint.WifiProfile is { } profile && !WifiProfilesFor(adapter).Contains(profile, StringComparer.Ordinal))
            return W(Word.ChooseAdapterWifiProfile);
        if (pair.OtherAdapterFor(endpoint) is null && !adapter.Enabled)
            return W(Word.LoadWifiProfiles);
        return null;
    }

    private static NetworkAdapterInfo? FindAdapter(NetworkAdapterInfo[] inventory, NetworkEndpoint? endpoint) =>
        endpoint is null ? null : inventory.FirstOrDefault(adapter => adapter.Id == endpoint.AdapterId);

    private string NameFor(NetworkEndpoint endpoint, NetworkAdapterInfo[] inventory) =>
        endpoint.WifiProfile ?? FindAdapter(inventory, endpoint)?.Name ?? W(Word.ConnectionUnavailable);

    private string[] WifiProfilesFor(NetworkAdapterInfo adapter)
    {
        var observed = wireless.FirstOrDefault(item => item.Id == adapter.Id);
        return observed is not null && (adapter.Enabled || observed.Profiles.Length > 0)
            ? observed.Profiles
            : settings.KnownWifiProfiles.GetValueOrDefault(adapter.Id) ?? [];
    }

    private void OpenAdapterWifiChoices(PairCard card, NetworkAdapterInfo adapter)
    {
        string[] profiles = WifiProfilesFor(adapter);
        var options = new List<DesktopChoice>
        {
            new(AutomaticWifiProfile, adapter.Name, W(Word.WifiAutomatic))
        };
        options.AddRange(profiles.Select(name => new DesktopChoice(WifiProfilePrefix + name, name)));
        if (administrator && !adapter.Enabled)
            options.Add(new DesktopChoice(LoadWifiNetworks, W(Word.LoadWifiProfiles)));

        var endpoint = card == first ? settings.NetworkOne : settings.NetworkTwo;
        string key = endpoint?.AdapterId == adapter.Id && endpoint.WifiProfile is { } selected
            ? WifiProfilePrefix + selected
            : AutomaticWifiProfile;
        card.Picker.OpenChoices(options, key, option =>
        {
            if (option == LoadWifiNetworks)
            {
                card.Picker.BeginInvoke(async () => await LoadAdapterWifiProfilesAsync(card, adapter.Id));
                return;
            }
            CommitAdapter(card, adapter, option == AutomaticWifiProfile ? null : option[WifiProfilePrefix.Length..]);
        });
    }

    private async Task LoadAdapterWifiProfilesAsync(PairCard card, Guid adapterId)
    {
        if (!administrator || busy || refreshing)
            return;

        SetBusy(true);
        try
        {
            await Task.Run(() => adapters.SetEnabled(adapterId, true));
            for (int attempt = 0; attempt < 10; attempt++)
            {
                await ReadAsync(true);
                if (wireless.Any(item => item.Id == adapterId && item.Profiles.Length > 0))
                    break;
                await Task.Delay(350);
            }
        }
        catch (Exception ex)
        {
            WriteJournal(Say("Не удалось включить адаптер Wi-Fi.", "Could not enable the Wi-Fi adapter."), ex.Message);
        }
        finally
        {
            await ReadAsync(false);
            lastRead = Environment.TickCount64;
            SetBusy(false);
        }

        if (ConnectionAdapters().FirstOrDefault(item => item.Id == adapterId) is { } adapter)
            OpenAdapterWifiChoices(card, adapter);
    }

    private string AdapterDetail(NetworkAdapterInfo? item) =>
        item is null ? "" : (item.Connected ? W(Word.Connected) : item.Enabled ?
            W(Word.EnabledWithoutLink) : W(Word.Disabled))
            + "   ·   " + item.Type + (item.IPv4 is { } address ? "   ·   " + address : "");

    private void SetAction(string? target, string placeholder)
    {
        action.Enabled = target is not null;
        action.Text = target is null ? placeholder : W(Word.Switch) + " " + target;
        tooltips.SetToolTip(action, target ?? "");
    }

    private void SetBusy(bool value)
    {
        busy = value;
        refresh.Enabled = preferences.Enabled = elevate.Enabled = filters.Enabled = !value;
        Render();
        if (!value && upgradeExitRequested)
            BeginInvoke(CloseForUpgradeWhenIdle);
    }

    private async Task RunConnectionAsync(NetworkEndpoint? requested, bool fromTray)
    {
        if (busy || refreshing)
            return;

        trayFeedback = null;
        SetBusy(true);
        if (fromTray)
            iconFeedback?.Begin(W(Word.Busy));
        bool success = false;
        string? message = null;
        string? diagnostic = null;
        try
        {
            await ReadConnectionsAsync();
            var inventory = ConnectionAdapters();
            var pair = Pair;
            var destination = requested ?? pair.ResolveTarget(inventory, wireless);
            if (destination is null)
            {
                message = PairProblem(pair, inventory);
                return;
            }

            message = AvailabilityProblem(destination, pair, inventory);
            if (message is not null)
                return;

            var otherAdapter = pair.OtherAdapterFor(destination);
            if (otherAdapter.HasValue && !administrator)
            {
                message = W(Word.ConnectionAdminRequired);
                if (fromTray && !tray.Visible)
                    notification.ShowBalloonTip(3000, "NetSwitch Pro", message, ToolTipIcon.Info);
                return;
            }

            var name = NameFor(destination, inventory);
            trayFeedback = W(Word.Switch) + " " + name + "…";
            if (fromTray)
                iconFeedback?.Begin(trayFeedback);
            UpdateTray();

            if (otherAdapter is { } source)
            {
                var outcome = await Task.Run(() => new AdapterSwitch(adapters, wifi: wifi)
                    .ApplyAsync(source, destination.AdapterId, destination.AdapterId, destination.WifiProfile));
                success = outcome.Result == AdapterSwitchResult.Changed;
                diagnostic = outcome.Diagnostic;
                message = outcome.Result switch
                {
                    AdapterSwitchResult.Changed => W(Word.Connected) + " · " + name,
                    AdapterSwitchResult.TargetNotReady => Say("Не удалось подтвердить готовность адаптера.", "Windows did not confirm the adapter was ready."),
                    AdapterSwitchResult.WifiProfileRequired => W(Word.ChooseAdapterWifiProfile),
                    AdapterSwitchResult.WifiProfileUnavailable => W(Word.ConnectionUnavailable),
                    AdapterSwitchResult.WifiNotConnected => Say("Wi-Fi не подключился. Исходное состояние восстановлено.", "Wi-Fi could not connect. The original adapter state was restored."),
                    AdapterSwitchResult.FailedRestored => Say("Переключение не удалось. Исходное состояние восстановлено.", "Switch failed. The original state was restored."),
                    _ => Say("Не удалось подтвердить восстановление адаптеров.", "Could not confirm that adapter state was restored.")
                };
            }
            else if (destination.WifiProfile is { } profile)
            {
                var outcome = await Task.Run(() => wifi.ConnectProfileAsync(destination.AdapterId, profile));
                success = outcome.Result == WifiSwitchResult.Connected;
                diagnostic = outcome.Diagnostic;
                message = outcome.Result switch
                {
                    WifiSwitchResult.Connected => W(Word.Connected) + " · " + name,
                    WifiSwitchResult.PreviousRestored => Say("Не удалось переключиться. Прежняя сеть восстановлена.", "Switch failed. The previous network was restored."),
                    WifiSwitchResult.ProfileUnavailable => W(Word.ConnectionUnavailable),
                    WifiSwitchResult.RestoreUnconfirmed => Say("Не удалось подтвердить восстановление прежней сети.", "Could not confirm restoration of the previous network."),
                    _ => Say("Windows не подтвердила подключение к ", "Windows did not confirm connection to ") + name
                };
            }
            else
                message = W(Word.PickConnections);
        }
        catch (Exception ex)
        {
            message = Say("Windows не выполнила переключение.", "Windows could not switch the connection.");
            diagnostic = ex.Message;
        }
        finally
        {
            if (message is not null)
                WriteJournal(message, diagnostic);
            trayFeedback = success ? null : message;
            await ReadConnectionsAsync();
            lastRead = Environment.TickCount64;
            SetBusy(false);
            if (fromTray)
                iconFeedback?.Complete(success, message ?? W(Word.Disconnect));
        }
    }

    private async Task OpenTrayAsync(Point location)
    {
        if (tray.Visible)
        {
            tray.Hide();
            return;
        }

        if (!busy)
            trayFeedback = null;
        UpdateSystemTheme();
        tray.OpenAt(location, TraySnapshot());
        await RefreshAsync(background: true);
    }

    private TrayState TraySnapshot()
    {
        var inventory = ConnectionAdapters();
        var pair = Pair;
        var active = pair.Current(inventory, wireless);
        var target = pair.ResolveTarget(inventory, wireless);
        TrayEntry Entry(NetworkEndpoint endpoint) => new(endpoint, NameFor(endpoint, inventory),
            ConnectionPair.Matches(active, endpoint), AvailabilityProblem(endpoint, pair, inventory) is null &&
                (pair.OtherAdapterFor(endpoint).HasValue || endpoint.WifiProfile is not null));
        var entries = new[] { pair.First, pair.Second }.OfType<NetworkEndpoint>().Distinct().Select(Entry).ToArray();
        var wifiAdapter = inventory.FirstOrDefault(item => item.Id == active?.AdapterId && item.IsWifi)
            ?? inventory.FirstOrDefault(item => pair.ContainsAdapter(item.Id) && item.IsWifi)
            ?? inventory.FirstOrDefault(item => item.IsWifi && item.Enabled);
        var saved = wifiAdapter is null ? [] : WifiProfilesFor(wifiAdapter)
            .Order(StringComparer.CurrentCultureIgnoreCase)
            .Select(profile => Entry(new(wifiAdapter.Id, profile))).ToArray();
        string? note = trayFeedback ?? (busy || refreshing ? W(Word.Busy) : null);
        return new TrayState(settings.Language, lightTheme, active is null ? W(Word.Disconnect) : NameFor(active, inventory),
            entries, saved, busy || refreshing, note,
            target is not null && AvailabilityProblem(target, pair, inventory) is null ? Entry(target) : null);
    }

    private void UpdateTray()
    {
        if (tray.Visible)
            tray.UpdateState(TraySnapshot());
    }

    internal void RestoreMainWindow()
    {
        tray.Hide();
        if (!Visible)
            Show();
        if (WindowState == FormWindowState.Minimized)
            WindowState = FormWindowState.Normal;
        Activate();
    }

    internal void ReportStartupError(Exception error) => WriteJournal(W(Word.StartupFailed), error.Message);

    internal void RequestUpgradeExit()
    {
        upgradeExitRequested = true;
        CloseForUpgradeWhenIdle();
    }

    private void CloseForUpgradeWhenIdle()
    {
        if (upgradeExitRequested && !busy && !refreshing && !IsDisposed)
        {
            activePreferences?.Close();
            ExitApplication();
        }
    }

    private void ExitApplication()
    {
        if (busy || refreshing)
            return;
        allowExit = true;
        Close();
    }

    protected override void Dispose(bool disposing)
    {
        bool disposeResources = disposing && !IsDisposed;
        if (disposeResources)
        {
            iconFeedback?.Dispose();
            notification.Dispose();
            tray.Dispose();
            wifi.Dispose();
            tooltips.Dispose();
        }
        base.Dispose(disposing);
        if (disposeResources)
            captionIcon.Dispose();
    }

    private void ChangeFilters()
    {
        if (binding)
            return;
        settings.ShowVirtual = showVirtual.Checked;
        settings.ShowBluetooth = showBluetooth.Checked;
        SaveAndRender();
    }

    private void SavePair(PairCard card)
    {
        if (!Guid.TryParse(card.Picker.SelectedKey, out var id) ||
            ConnectionAdapters().FirstOrDefault(item => item.Id == id) is not { } adapter)
            return;

        var previous = card == first ? settings.NetworkOne : settings.NetworkTwo;
        if (previous?.AdapterId != id || !adapter.IsWifi)
            CommitAdapter(card, adapter, null);
        else
            Render();

        if (adapter.IsWifi)
            card.Picker.BeginInvoke(() => OpenAdapterWifiChoices(card, adapter));
    }

    private void CommitAdapter(PairCard card, NetworkAdapterInfo adapter, string? wifiProfile)
    {
        var endpoint = new NetworkEndpoint(adapter.Id, wifiProfile);
        if (card == first)
            settings.NetworkOne = endpoint;
        else
            settings.NetworkTwo = endpoint;
        settings.UnifiedPairSaved = true;
        SaveAndRender();
    }

    private void SaveAndRender()
    {
        SaveSettings();
        Render();
    }

    private void SaveSettings()
    {
        try { settings.Save(); }
        catch (Exception ex) { WriteJournal(Say("Не удалось сохранить настройки.", "Could not save settings."), ex.Message); }
    }

    private void Elevate()
    {
        if (busy || refreshing)
            return;
        SaveSettings();
        try
        {
            using var child = Process.Start(new ProcessStartInfo(Application.ExecutablePath)
            {
                UseShellExecute = true,
                Verb = "runas"
            });
            if (child is not null)
            {
                allowExit = true;
                Close();
            }
        }
        catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223) { }
        catch (Exception ex) { WriteJournal(Say("Не удалось запустить режим администратора.", "Could not start administrator mode."), ex.Message); }
    }

    private void WriteJournal(string message, string? diagnostic = null, bool showStatus = true)
    {
        journal.Insert(0, $"{DateTime.Now:HH:mm:ss}  {message}");
        if (diagnostic is not null)
            journal.Insert(1, "    " + diagnostic);
        if (journal.Count > 100)
            journal.RemoveRange(100, journal.Count - 100);
        if (showStatus)
        {
            status.Text = journal[0];
            ShowFeedback();
        }
    }

    private void ShowFeedback()
    {
        status.Visible = status.Text.Length > 0;
        hint.Visible = !status.Visible;
        tooltips.SetToolTip(status, status.Text);
        tooltips.SetToolTip(hint, hint.Text);
    }

    private void ShowJournal()
    {
        bool empty = journal.Count == 0;
        string body = empty ? W(Word.EmptyLog) : string.Join(Environment.NewLine, journal);
        int height = empty ? 78 : Math.Clamp(58 + journal.Count * 25, 130, 316);
        using var dialog = Dialog(W(Word.Log), new Size(empty ? 365 : 540, height));
        dialog.Padding = new Padding(9);

        var content = Grid(2, empty ? 1 : 2);
        content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 142));
        content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var activity = new JournalView(body) { Dock = DockStyle.Fill, Margin = Padding.Empty };
        var panel = Surface(1, 1);
        panel.Padding = new Padding(8, 5, 8, 5);
        panel.Controls.Add(activity, 0, 0);
        content.Controls.Add(panel, 0, 0);
        content.SetColumnSpan(panel, 2);
        if (!empty)
        {
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 39));
            var copy = MakeButton(ButtonAppearance.Outline);
            copy.Text = W(Word.CopyLog);
            copy.Margin = new Padding(0, 6, 0, 0);
            copy.Click += (_, _) =>
            {
                try { Clipboard.SetText(body); }
                catch (ExternalException) { return; }
                copy.Text = W(Word.LogCopied);
            };
            content.Controls.Add(copy, 1, 1);
        }

        dialog.Controls.Add(content);
        PaintTheme(dialog);
        dialog.HandleCreated += (_, _) => SetTitleBar(dialog);
        dialog.ResumeLayout(true);
        dialog.ShowDialog(this);
    }

    private void ShowPreferences()
    {
        using var dialog = Dialog(W(Word.Settings), new Size(504, 320));
        var layout = Grid(2, 5);
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Insert(2, new RowStyle(SizeType.Absolute, 53));
        layout.RowStyles.Insert(3, new RowStyle(SizeType.Absolute, 78));
        var theme = MakePicker();
        theme.Margin = new Padding(0, 2, 0, 4);
        Fill(theme, [
            new DesktopChoice("system", W(Word.SystemTheme)),
            new DesktopChoice("dark", W(Word.Dark)),
            new DesktopChoice("light", W(Word.Light))], settings.Theme);
        theme.AccessibleName = W(Word.Theme);
        var languages = MakePicker();
        languages.Margin = theme.Margin;
        languages.AccessibleName = W(Word.Language);
        Fill(languages, UiWords.Languages.Select(x => new DesktopChoice(x.Key, x.Label)), settings.Language);
        var apply = MakeButton(ButtonAppearance.Primary);
        apply.Text = Say("Сохранить", "Save");
        apply.DialogResult = DialogResult.OK;
        layout.Controls.Add(new Label { Text = W(Word.Theme), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
        layout.Controls.Add(theme, 1, 0);
        layout.Controls.Add(new Label { Text = W(Word.Language), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
        layout.Controls.Add(languages, 1, 1);
        var startup = new DesktopToggle { Text = W(Word.Autostart), Dock = DockStyle.Fill,
            SwitchStyle = true, Margin = new Padding(0, 6, 0, 3) };
        string? startupCommand = null;
        try
        {
            startupCommand = StartupRegistration.ReadCommand();
            startup.Checked = !string.IsNullOrWhiteSpace(startupCommand);
        }
        catch (Exception error) when (IsStartupError(error))
        {
            startup.Enabled = false;
            WriteJournal(W(Word.StartupFailed), error.Message);
        }
        layout.Controls.Add(startup, 0, 2);
        layout.SetColumnSpan(startup, 2);
        AddUpdateControls(dialog, layout);
        layout.Controls.Add(apply, 1, 4);
        apply.Margin = new Padding(0, 9, 0, 0);
        dialog.AcceptButton = apply;
        dialog.Controls.Add(layout);
        PaintTheme(dialog);
        dialog.HandleCreated += (_, _) => SetTitleBar(dialog);
        dialog.ResumeLayout(true);
        activePreferences = dialog;
        DialogResult result;
        try { result = dialog.ShowDialog(this); }
        finally { activePreferences = null; }
        if (result != DialogResult.OK)
            return;
        if (startup.Enabled)
        {
            try
            {
                bool registered = !string.IsNullOrWhiteSpace(startupCommand);
                if (startup.Checked != registered)
                    StartupRegistration.SetEnabled(startup.Checked, Application.ExecutablePath);
            }
            catch (Exception error) when (IsStartupError(error))
            {
                WriteJournal(W(Word.StartupFailed), error.Message);
                MessageBox.Show(this, W(Word.StartupFailed), "NetSwitch Pro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        settings.Theme = theme.SelectedKey ?? "system";
        lightTheme = settings.IsLight;
        settings.Language = languages.SelectedKey ?? "English";
        UpdateLabels();
        Render();
        PaintTheme(this);
        SaveAndRender();
    }

    private void AddUpdateControls(Form dialog, TableLayoutPanel layout)
    {
        var row = Grid(2, 1);
        row.Margin = new Padding(0, 8, 0, 0);
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 205));
        row.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var version = typeof(MainForm).Assembly.GetName().Version!;
        var text = TextLabel(10, bold: true);
        text.Text = Say("Версия ", "Version ") + version.ToString(3);
        var note = TextLabel(9, "muted");
        var labels = Grid(1, 2);
        labels.RowStyles.Add(new RowStyle(SizeType.Absolute, 27));
        labels.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        labels.Controls.Add(text, 0, 0);
        labels.Controls.Add(note, 0, 1);
        labels.Margin = new Padding(0, 0, 10, 0);
        row.Controls.Add(labels, 0, 0);

        var actions = Grid(1, 2);
        actions.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        actions.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        var check = MakeButton(ButtonAppearance.Outline);
        var manual = MakeButton();
        var checkLabel = Say("Проверить обновления", "Check for updates");
        check.Text = checkLabel;
        manual.Text = Say("Скачать вручную", "Download manually");
        manual.Visible = false;
        actions.Controls.Add(check, 0, 0);
        actions.Controls.Add(manual, 0, 1);
        row.Controls.Add(actions, 1, 0);
        layout.Controls.Add(row, 0, 3);
        layout.SetColumnSpan(row, 2);

        var lifetime = new CancellationTokenSource();
        dialog.FormClosed += (_, _) => lifetime.Cancel();
        dialog.Disposed += (_, _) => lifetime.Dispose();
        var token = lifetime.Token;
        AvailableUpdate? available = null;

        manual.Click += (_, _) =>
        {
            try
            {
                using var browser = Process.Start(new ProcessStartInfo(ReleaseUpdates.ReleasesPage) { UseShellExecute = true });
            }
            catch (Exception error) when (error is System.ComponentModel.Win32Exception or InvalidOperationException)
            {
                note.Text = Say("Не удалось открыть GitHub.", "Could not open GitHub.");
            }
        };
        check.Click += async (_, _) =>
        {
            check.Enabled = false;
            manual.Visible = false;
            try
            {
                if (available is null)
                {
                    note.Text = Say("Проверка обновлений…", "Checking for updates…");
                    available = await ReleaseUpdates.CheckAsync(version, token);
                    if (dialog.IsDisposed)
                        return;
                    note.Text = available is null ? Say("Установлена актуальная версия.", "NetSwitch Pro is up to date.") :
                        Say("Доступна версия ", "Available version ") + available.Version.ToString(3);
                    if (available is not null && !await Task.Run(() => UpdatePublisher.IsTrusted(Application.ExecutablePath), token))
                    {
                        if (dialog.IsDisposed)
                            return;
                        note.Text = Say("Доступна версия " + available.Version.ToString(3) + ". Требуется ручная загрузка.",
                            "Version " + available.Version.ToString(3) + " is available for manual download.");
                        manual.Visible = true;
                        available = null;
                    }
                    if (!dialog.IsDisposed)
                        check.Text = available is null ? checkLabel : Say("Обновить", "Update");
                    return;
                }

                note.Text = Say("Подготовка загрузки…", "Preparing download…");
                var progress = new Progress<int>(percent =>
                {
                    if (!dialog.IsDisposed)
                        note.Text = Say("Загрузка ", "Downloading ") + percent + "%";
                });
                string executable = await ReleaseUpdates.DownloadAsync(available, progress, token);
                token.ThrowIfCancellationRequested();
                note.Text = Say("Завершение обновления…", "Finishing update…");
                using var child = Process.Start(new ProcessStartInfo(executable, ReleaseUpdates.InstallArgument) { UseShellExecute = false })
                    ?? throw new IOException("Windows did not start the update.");
                await child.WaitForExitAsync(token);
                if (!token.IsCancellationRequested)
                    throw new IOException("The update closed before the current application could finish its handoff.");
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { }
            catch (Exception error) when (error is IOException or InvalidDataException or System.Net.Http.HttpRequestException or System.Text.Json.JsonException
                or UnauthorizedAccessException or System.ComponentModel.Win32Exception or OperationCanceledException)
            {
                if (!dialog.IsDisposed)
                {
                    note.Text = Say("Не удалось обновить. См. журнал.", "Update failed. See activity.");
                    manual.Visible = true;
                    available = null;
                    check.Text = checkLabel;
                    WriteJournal(note.Text, error.Message, showStatus: false);
                }
            }
            finally
            {
                if (!dialog.IsDisposed)
                    check.Enabled = true;
            }
        };
    }

    private static bool IsStartupError(Exception error) =>
        error is IOException or UnauthorizedAccessException or System.Security.SecurityException;

    private Form Dialog(string title, Size size)
    {
        var dialog = new Form();
        dialog.SuspendLayout();
        dialog.AutoScaleDimensions = new SizeF(96, 96);
        dialog.AutoScaleMode = AutoScaleMode.Dpi;
        dialog.Font = Font;
        dialog.Text = title;
        dialog.Padding = new Padding(20);
        dialog.ClientSize = size;
        dialog.Icon = Icon;
        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
        dialog.StartPosition = FormStartPosition.CenterParent;
        dialog.ShowInTaskbar = dialog.MinimizeBox = dialog.MaximizeBox = false;
        return dialog;
    }

    private void PaintTheme(Control control)
    {
        var palette = DesktopPalette.Resolve(settings.IsLight);
        control.BackColor = control switch
        {
            Form => palette.Window,
            SurfaceGrid or TextBox => palette.Surface,
            _ when Equals(control.Tag, "line") => palette.Line,
            _ => control.Parent?.BackColor ?? palette.Window
        };
        control.ForeColor = control switch
        {
            SurfaceGrid => palette.Line,
            _ when Equals(control.Tag, "muted") => palette.Muted,
            _ => palette.Ink
        };
        if (control is DesktopButton button)
            button.Palette = palette;
        if (control is DesktopToggle toggle)
            toggle.Palette = palette;
        if (control is JournalView activity)
            activity.Palette = palette;
        if (control is Form form && form.IsHandleCreated)
            SetTitleBar(form);
        foreach (Control child in control.Controls)
            PaintTheme(child);
        control.Invalidate();
    }

    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);
        if (m.Msg is 0x001A or 0x031A)
            UpdateSystemTheme();
    }

    private void UpdateSystemTheme()
    {
        if (!hasShown || settings.Theme != "system")
            return;
        bool next = settings.IsLight;
        if (next == lightTheme)
            return;

        lightTheme = next;
        PaintTheme(this);
        foreach (var dialog in OwnedForms)
            PaintTheme(dialog);
        UpdateTray();
    }

    private void SetTitleBar(Form form)
    {
        const int UseImmersiveDarkMode = 20, CaptionColor = 35, CaptionTextColor = 36;
        const int SystemDefaultColor = -1;
        bool highContrast = SystemInformation.HighContrast;
        var palette = DesktopPalette.Resolve(lightTheme);
        int dark = !lightTheme && !highContrast ? 1 : 0;
        int background = highContrast ? SystemDefaultColor : ColorTranslator.ToWin32(palette.Window);
        int text = highContrast ? SystemDefaultColor : ColorTranslator.ToWin32(palette.Ink);
        _ = DwmSetWindowAttribute(form.Handle, UseImmersiveDarkMode, ref dark, sizeof(int));
        _ = DwmSetWindowAttribute(form.Handle, CaptionColor, ref background, sizeof(int));
        _ = DwmSetWindowAttribute(form.Handle, CaptionTextColor, ref text, sizeof(int));
    }

    private static bool IsAdministrator()
    {
        using var identity = WindowsIdentity.GetCurrent();
        return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static TableLayoutPanel Grid(int columns, int rows) => new LayoutGrid()
    {
        Dock = DockStyle.Fill, Margin = Padding.Empty, ColumnCount = columns, RowCount = rows
    };

    private static Label TextLabel(float size, string? role = null, bool bold = false) => new()
    {
        Dock = DockStyle.Fill, Margin = Padding.Empty, AutoEllipsis = true, Tag = role,
        TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular)
    };

    private static TableLayoutPanel Surface(int columns, int rows) => new SurfaceGrid
    {
        Dock = DockStyle.Fill, Margin = Padding.Empty, ColumnCount = columns, RowCount = rows
    };

    private static DesktopButton MakeButton(ButtonAppearance appearance = ButtonAppearance.Quiet, string? glyph = null) => new()
    {
        Dock = DockStyle.Fill, Margin = Padding.Empty, Appearance = appearance, Glyph = glyph
    };

    private static DesktopPicker MakePicker() => new()
    {
        Dock = DockStyle.Fill, Margin = Padding.Empty
    };

    private static void Fill(DesktopPicker picker, IEnumerable<DesktopChoice> choices, string? selected) => picker.Bind(choices, selected);

    private sealed class PairCard
    {
        public TableLayoutPanel Container { get; } = Grid(2, 1);
        public Label Caption { get; } = TextLabel(11, bold: true);
        public DesktopPicker Picker { get; } = MakePicker();
        public Label Detail { get; } = TextLabel(9, "muted");

        public PairCard()
        {
            Container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            Container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            Container.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var labels = Grid(1, 2);
            labels.Padding = new Padding(0, 6, 8, 6);
            labels.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            labels.RowStyles.Add(new RowStyle(SizeType.Percent, 0));
            labels.Controls.Add(Caption, 0, 0);
            labels.Controls.Add(Detail, 0, 1);
            Detail.Visible = false;
            Detail.TextChanged += (_, _) =>
            {
                bool shown = Detail.Text.Length > 0;
                labels.RowStyles[0].Height = shown ? 56 : 100;
                labels.RowStyles[1].Height = shown ? 44 : 0;
                Detail.Visible = shown;
            };
            Picker.Margin = new Padding(0, 10, 0, 10);
            Container.Controls.Add(labels, 0, 0);
            Container.Controls.Add(Picker, 1, 0);
            Caption.TextChanged += (_, _) => Picker.AccessibleName = Caption.Text;
        }
    }

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);
}
