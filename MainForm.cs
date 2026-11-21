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
    private readonly Label interfaceLabel = TextLabel(10, "muted");
    private readonly Label activeLabel = TextLabel(10, "muted");
    private readonly Label current = TextLabel(19, bold: true);
    private readonly Label signal = TextLabel(9, "muted");
    private readonly Label pairTitle = TextLabel(13, bold: true);
    private readonly Label accessTitle = TextLabel(10, "muted");
    private readonly Label accessValue = TextLabel(15, bold: true);
    private readonly Label accessDetail = TextLabel(9, "muted");
    private readonly Label hint = TextLabel(10, "muted");
    private readonly Label status = TextLabel(10, "muted");
    private readonly DesktopButton wifiTab = MakeButton(ButtonAppearance.Segment, "");
    private readonly DesktopButton adapterTab = MakeButton(ButtonAppearance.Segment, "");
    private readonly DesktopButton refresh = MakeButton(glyph: "");
    private readonly DesktopButton history = MakeButton(glyph: "");
    private readonly DesktopButton preferences = MakeButton(glyph: "");
    private readonly DesktopButton elevate = MakeButton(ButtonAppearance.Outline, glyph: "\uE72E");
    private readonly DesktopButton action = MakeButton(ButtonAppearance.Primary, "");
    private readonly DesktopPicker wifiInterface = MakePicker();
    private readonly DesktopToggle showVirtual = new();
    private readonly DesktopToggle showBluetooth = new();
    private readonly PairCard first = new();
    private readonly PairCard second = new();
    private readonly TableLayoutPanel connection = Surface(2, 2);
    private readonly TableLayoutPanel adapterAccess = Surface(2, 1);
    private readonly FlowLayoutPanel filters = new() { Dock = DockStyle.Fill, WrapContents = true, Margin = Padding.Empty };
    private readonly ToolTip tooltips = new();
    private IReadOnlyList<WifiAdapter> wireless = [];
    private IReadOnlyList<NetworkAdapterInfo> wired = [];
    private bool wifiMode;
    private bool busy;
    private bool hasShown;
    private bool binding;
    private string? readError;

    public MainForm(bool forceAdapters)
    {
        wifiMode = !forceAdapters && !settings.ShowAdapters;
        SuspendLayout();
        Text = "NetSwitch Pro";
        AutoScaleDimensions = new SizeF(96, 96);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 10);
        ClientSize = new Size(656, 507);
        MinimumSize = new Size(607, 543);
        StartPosition = FormStartPosition.CenterScreen;
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        BuildLayout();
        HandleCreated += (_, _) => SetTitleBar(this);
        UpdateLabels();
        Render();
        PaintTheme(this);
        ResumeLayout(true);

        wifiTab.Click += async (_, _) => await SetModeAsync(true);
        adapterTab.Click += async (_, _) => await SetModeAsync(false);
        refresh.Click += async (_, _) => await RefreshAsync();
        history.Click += (_, _) => ShowJournal();
        preferences.Click += (_, _) => ShowPreferences();
        elevate.Click += (_, _) => Elevate();
        action.Click += async (_, _) => await SwitchAsync();
        wifiInterface.SelectionChangeCommitted += (_, _) => {
            if (Selected(wifiInterface) is { } id)
                settings.WifiAdapterId = Guid.Parse(id);
            SaveAndRender();
        };
        first.Picker.SelectionChangeCommitted += (_, _) => SavePair(first);
        second.Picker.SelectionChangeCommitted += (_, _) => SavePair(second);
        showVirtual.CheckedChanged += (_, _) => ChangeFilters();
        showBluetooth.CheckedChanged += (_, _) => ChangeFilters();
        Shown += async (_, _) =>
        {
            hasShown = true;
            await RefreshAsync();
        };
        Activated += async (_, _) =>
        {
            if (hasShown)
                await RefreshAsync();
        };
        FormClosing += (_, e) => e.Cancel = busy && e.CloseReason == CloseReason.UserClosing;
        FormClosed += (_, _) => { wifi.Dispose(); tooltips.Dispose(); };
    }

    private void BuildLayout()
    {
        var root = Grid(1, 7);
        root.Padding = new Padding(24, 19, 24, 16);
        foreach (int height in new[] { 61, 57, 104, 41, 134 })
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

        var tabs = Surface(2, 1);
        tabs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        tabs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        tabs.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        tabs.Padding = new Padding(4);
        tabs.Margin = new Padding(0, 0, 0, 13);
        tabs.Controls.Add(wifiTab, 0, 0);
        tabs.Controls.Add(adapterTab, 1, 0);
        root.Controls.Add(tabs, 0, 1);

        connection.Padding = new Padding(15, 12, 15, 10);
        connection.Margin = new Padding(0, 0, 0, 12);
        connection.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 46));
        connection.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 54));
        connection.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
        connection.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        connection.Controls.Add(activeLabel, 0, 0);
        connection.Controls.Add(interfaceLabel, 1, 0);
        connection.Controls.Add(wifiInterface, 1, 1);

        signal.Text = "●";
        var activeValue = Grid(2, 1);
        activeValue.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 18));
        activeValue.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        activeValue.Controls.Add(signal, 0, 0);
        activeValue.Controls.Add(current, 1, 0);
        activeValue.Margin = new Padding(0, 0, 12, 0);
        connection.Controls.Add(activeValue, 0, 1);

        var connectionHost = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 12) };
        connection.Margin = Padding.Empty;
        connectionHost.Controls.Add(connection);
        connectionHost.Controls.Add(adapterAccess);
        BuildAdapterAccess();
        root.Controls.Add(connectionHost, 0, 2);

        var pairHeader = Grid(2, 1);
        pairHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        pairHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 248));
        pairHeader.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        pairHeader.Controls.Add(pairTitle, 0, 0);
        filters.FlowDirection = FlowDirection.RightToLeft;
        filters.WrapContents = false;
        filters.Padding = new Padding(0, 3, 0, 0);
        showVirtual.Width = 132;
        showBluetooth.Width = 104;
        filters.Controls.Add(showBluetooth);
        filters.Controls.Add(showVirtual);
        pairHeader.Controls.Add(filters, 1, 0);
        root.Controls.Add(pairHeader, 0, 3);

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
        footer.Controls.Add(action, 1, 0);
        root.Controls.Add(footer, 0, 6);
    }

    private void BuildAdapterAccess()
    {
        adapterAccess.Padding = new Padding(15, 10, 15, 10);
        adapterAccess.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62));
        adapterAccess.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
        adapterAccess.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var captions = Grid(1, 3);
        captions.RowStyles.Add(new RowStyle(SizeType.Percent, 29));
        captions.RowStyles.Add(new RowStyle(SizeType.Percent, 42));
        captions.RowStyles.Add(new RowStyle(SizeType.Percent, 29));
        captions.Controls.Add(accessTitle, 0, 0);
        captions.Controls.Add(accessValue, 0, 1);
        captions.Controls.Add(accessDetail, 0, 2);
        adapterAccess.Controls.Add(captions, 0, 0);

        elevate.Margin = new Padding(12, 15, 0, 15);
        adapterAccess.Controls.Add(elevate, 1, 0);
    }

    private string W(Word word) => UiWords.Get(settings.Language, word);
    private string Say(string ru, string en) => settings.Language == "Russian" ? ru : en;

    private void UpdateLabels()
    {
        wifiTab.Text = W(Word.Wifi);
        adapterTab.Text = W(Word.Adapters);
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

    private async Task SetModeAsync(bool showWifi)
    {
        if (busy || wifiMode == showWifi)
            return;
        wifiMode = showWifi;
        settings.ShowAdapters = !showWifi;
        SaveSettings();
        await RefreshAsync();
    }

    private async Task RefreshAsync()
    {
        if (busy)
            return;
        SetBusy(true);
        try { await ReadAsync(); }
        finally { SetBusy(false); }
    }

    private async Task ReadAsync()
    {
        readError = null;
        try
        {
            if (wifiMode)
                wireless = await Task.Run(() => wifi.GetAdapters());
            else
                wired = await Task.Run(adapters.GetAdapters);
        }
        catch (Exception ex)
        {
            if (wifiMode)
                wireless = [];
            else
                wired = [];
            readError = Say("Не удалось прочитать данные Windows. См. журнал.", "Windows data is unavailable. See activity.");
            WriteJournal(ex.Message, showStatus: false);
        }
    }

    private WifiAdapter? ActiveInterface()
    {
        return wireless.FirstOrDefault(x => x.Id == settings.WifiAdapterId)
            ?? wireless.FirstOrDefault(x => x.Connected)
            ?? wireless.FirstOrDefault();
    }

    private IEnumerable<NetworkAdapterInfo> VisibleAdapters() => wired
        .Where(x => (settings.ShowVirtual || !x.IsVirtual) && (settings.ShowBluetooth || !x.IsBluetooth));

    private (NetworkAdapterInfo? First, NetworkAdapterInfo? Second) VisiblePair()
    {
        var rows = VisibleAdapters().ToArray();
        return (Resolve(rows, settings.AdapterAId, settings.AdapterA), Resolve(rows, settings.AdapterBId, settings.AdapterB));
    }

    private static NetworkAdapterInfo? Resolve(NetworkAdapterInfo[] rows, Guid? id, string? name) =>
        id is { } value ? rows.FirstOrDefault(x => x.Id == value) :
        rows.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    private void Render()
    {
        binding = true;
        SuspendLayout();
        connection.SuspendLayout();
        try
        {
            wifiTab.Selected = wifiMode;
            adapterTab.Selected = !wifiMode;
            wifiTab.Invalidate();
            adapterTab.Invalidate();
            first.Picker.Enabled = second.Picker.Enabled = !busy;
            status.Text = readError ?? (journal.Count > 0 ? journal[0] : "");
            if (wifiMode)
                RenderWifi();
            else
                RenderAdapters();
            if (busy)
            {
                action.Text = W(Word.Busy);
                action.Enabled = false;
            }
            ShowFeedback();
            action.Invalidate();
        }
        finally
        {
            connection.ResumeLayout(true);
            ResumeLayout(true);
            binding = false;
        }
    }

    private void RenderWifi()
    {
        filters.Visible = adapterAccess.Visible = false;
        connection.Visible = true;
        pairTitle.Text = Say("Пара Wi-Fi-сетей", "Wi-Fi pair");
        wifiInterface.AccessibleName = W(Word.Interface);
        interfaceLabel.Text = W(Word.Interface);
        activeLabel.Text = W(Word.Active);
        first.Caption.Text = W(Word.FirstNetwork);
        second.Caption.Text = W(Word.SecondNetwork);
        first.Picker.Placeholder = second.Picker.Placeholder =
            Say("Выберите сеть", "Choose a network");
        var selected = ActiveInterface();
        wifiInterface.Placeholder = Say("Нет адаптера", "No adapter");
        Fill(wifiInterface, wireless.Select(x => new DesktopChoice(x.Id.ToString(), x.Description)), selected?.Id.ToString());
        var profiles = selected?.Profiles ?? [];
        Fill(first.Picker, profiles.Select(x => new DesktopChoice(x, x)), settings.WifiA);
        Fill(second.Picker, profiles.Select(x => new DesktopChoice(x, x)), settings.WifiB);
        string? a = Selected(first.Picker), b = Selected(second.Picker);
        current.Text = selected?.ActiveProfile ?? (selected?.Connected == true ? W(Word.Unknown) : W(Word.Disconnect));
        SetSignal(selected?.Connected == true);
        tooltips.SetToolTip(current, current.Text);
        tooltips.SetToolTip(wifiInterface, selected?.Description ?? "");
        first.Detail.Text = WifiDetail(a, selected?.ActiveProfile);
        second.Detail.Text = WifiDetail(b, selected?.ActiveProfile);
        tooltips.SetToolTip(first.Detail, "");
        tooltips.SetToolTip(second.Detail, "");
        string? target = WifiService.ChooseTarget(a, b, selected?.ActiveProfile);
        hint.Text = selected is null
            ? Say("Windows не обнаружила Wi-Fi-адаптер.", "Windows did not detect a Wi-Fi adapter.")
            : profiles.Length < 2
                ? Say("Для пары нужны две сети, сохранённые в настройках Wi-Fi Windows.", "Save two Wi-Fi networks in Windows first.")
                : target is null ? W(Word.PickWifi) : "";
        SetAction(target, W(Word.Connect));
    }

    private string WifiDetail(string? selected, string? active) =>
        selected is null ? "" : selected == active ? W(Word.Connected) : W(Word.Saved);

    private void RenderAdapters()
    {
        connection.Visible = false;
        adapterAccess.Visible = filters.Visible = true;
        elevate.Visible = !administrator;
        pairTitle.Text = Say("Пара адаптеров", "Adapter pair");
        accessTitle.Text = Say("Управление адаптерами", "Adapter control");
        accessValue.Text = administrator ? Say("Управление разрешено", "Control allowed") : Say("Режим просмотра", "View only");
        accessDetail.Text = administrator
            ? Say("Адаптеры можно включать и отключать", "You can enable and disable adapters")
            : Say("Для переключения нужны права администратора", "Switching requires administrator rights");
        showVirtual.Checked = settings.ShowVirtual;
        showBluetooth.Checked = settings.ShowBluetooth;
        first.Caption.Text = W(Word.FirstAdapter);
        second.Caption.Text = W(Word.SecondAdapter);
        first.Picker.Placeholder = second.Picker.Placeholder =
            Say("Выберите адаптер", "Choose an adapter");
        var (a, b) = VisiblePair();
        var choices = VisibleAdapters().Select(x => new DesktopChoice(
            x.Id.ToString(), x.Name, AdapterDetail(x), x.Connected)).ToArray();
        Fill(first.Picker, choices, a?.Id.ToString());
        Fill(second.Picker, choices, b?.Id.ToString());
        first.Detail.Text = AdapterDetail(a);
        second.Detail.Text = AdapterDetail(b);
        tooltips.SetToolTip(first.Detail, a?.Description ?? "");
        tooltips.SetToolTip(second.Detail, b?.Description ?? "");
        var target = a is not null && b is not null && a.Id != b.Id ? AdapterSwitch.ChooseTarget(a, b).Name : null;
        hint.Text = target is null ? W(Word.PickAdapters) : W(Word.AdapterHint);
        SetAction(target, W(Word.Switch));
        action.Enabled = administrator && target is not null;
        if (target is null)
            action.Text = Say("Переключить адаптер", "Switch adapter");
    }

    private void SetSignal(bool active)
    {
        signal.Tag = active ? "signal" : "muted";
        signal.ForeColor = active ? DesktopPalette.Resolve(settings.IsLight).Signal : DesktopPalette.Resolve(settings.IsLight).Muted;
    }

    private string AdapterDetail(NetworkAdapterInfo? item) =>
        item is null ? "" : (item.Connected ? W(Word.Connected) : item.Enabled ? W(Word.Enabled) : W(Word.Disabled))
            + "   ·   " + item.Type + (item.IPv4 is { } address ? "   ·   " + address : "");

    private void SetAction(string? target, string verb)
    {
        action.Enabled = target is not null;
        action.Text = target is null ? verb : verb + " " + target;
        tooltips.SetToolTip(action, target ?? "");
    }

    private void SetBusy(bool value)
    {
        busy = value;
        wifiTab.Enabled = adapterTab.Enabled = refresh.Enabled = preferences.Enabled = !value;
        wifiInterface.Enabled = elevate.Enabled = filters.Enabled = !value;
        Render();
    }

    private async Task SwitchAsync()
    {
        if (busy || !action.Enabled)
            return;
        var selected = ActiveInterface();
        string? firstSelection = Selected(first.Picker), secondSelection = Selected(second.Picker);
        var (a, b) = VisiblePair();
        SetBusy(true);
        try
        {
            if (wifiMode && selected is not null && firstSelection is not null && secondSelection is not null &&
                WifiService.ChooseTarget(firstSelection, secondSelection, selected.ActiveProfile) is { } target)
            {
                var outcome = await Task.Run(() => wifi.SwitchAsync(selected.Id, firstSelection, secondSelection, target));
                var message = outcome.Result switch
                {
                    WifiSwitchResult.Connected => Say("Подключено к ", "Connected to ") + target,
                    WifiSwitchResult.PreviousRestored => Say("Не удалось переключиться. Прежняя сеть восстановлена.", "Switch failed. The previous network was restored."),
                    WifiSwitchResult.RestoreUnconfirmed => Say("Не удалось подтвердить восстановление прежней сети.", "Could not confirm restoration of the previous network."),
                    _ => Say("Windows не подтвердила подключение к ", "Windows did not confirm connection to ") + target
                };
                WriteJournal(message, outcome.Diagnostic);
            }
            else if (!wifiMode && administrator && a is not null && b is not null && a.Id != b.Id)
            {
                var outcome = await Task.Run(() => new AdapterSwitch(adapters).ApplyAsync(a.Id, b.Id));
                var message = outcome.Result switch
                {
                    AdapterSwitchResult.Changed => Say("Адаптер переключён.", "Adapter switched."),
                    AdapterSwitchResult.TargetNotReady => Say("Второй адаптер не готов к переключению.", "The other adapter was not ready."),
                    AdapterSwitchResult.FailedRestored => Say("Переключение не удалось. Исходное состояние восстановлено.", "Switch failed. The original state was restored."),
                    _ => Say("Не удалось подтвердить восстановление адаптеров.", "Could not confirm that adapter state was restored.")
                };
                WriteJournal(message, outcome.Diagnostic);
            }
        }
        catch (Exception ex)
        {
            WriteJournal(Say("Windows не выполнила переключение.", "Windows could not switch the connection."), ex.Message);
        }
        finally
        {
            await ReadAsync();
            SetBusy(false);
        }
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
        var key = Selected(card.Picker);
        if (wifiMode)
        {
            if (card == first)
                settings.WifiA = key;
            else
                settings.WifiB = key;
        }
        else if (Guid.TryParse(key, out var id))
        {
            if (card == first)
            {
                settings.AdapterAId = id;
                settings.AdapterA = null;
            }
            else
            {
                settings.AdapterBId = id;
                settings.AdapterB = null;
            }
        }
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
        SaveSettings();
        try
        {
            using var child = Process.Start(new ProcessStartInfo(Application.ExecutablePath, "--adapters")
            {
                UseShellExecute = true,
                Verb = "runas"
            });
            if (child is not null)
                Close();
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
        using var dialog = Dialog(W(Word.Log), new Size(630, 340));
        var text = new TextBox
        {
            Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill, BorderStyle = BorderStyle.None,
            Text = journal.Count == 0 ? Say("Событий пока нет.", "No activity yet.") : string.Join(Environment.NewLine, journal)
        };
        var panel = Surface(1, 1);
        panel.Padding = new Padding(13);
        panel.Controls.Add(text, 0, 0);
        dialog.Controls.Add(panel);
        PaintTheme(dialog);
        dialog.HandleCreated += (_, _) => SetTitleBar(dialog);
        dialog.ResumeLayout(true);
        dialog.ShowDialog(this);
    }

    private void ShowPreferences()
    {
        using var dialog = Dialog(W(Word.Settings), new Size(410, 183));
        var layout = Grid(2, 3);
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var theme = MakePicker();
        theme.Margin = new Padding(0, 2, 0, 4);
        Fill(theme, [new DesktopChoice("dark", W(Word.Dark)), new DesktopChoice("light", W(Word.Light))], settings.IsLight ? "light" : "dark");
        theme.AccessibleName = W(Word.Theme);
        var languages = MakePicker();
        languages.Margin = theme.Margin;
        languages.AccessibleName = W(Word.Language);
        Fill(languages, UiWords.Languages.Select(x => new DesktopChoice(x.Key, x.Label)), settings.Language);
        var apply = MakeButton(ButtonAppearance.Confirm);
        apply.Text = Say("Сохранить", "Save");
        apply.DialogResult = DialogResult.OK;
        layout.Controls.Add(new Label { Text = W(Word.Theme), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
        layout.Controls.Add(theme, 1, 0);
        layout.Controls.Add(new Label { Text = W(Word.Language), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
        layout.Controls.Add(languages, 1, 1);
        layout.Controls.Add(apply, 1, 2);
        apply.Margin = new Padding(0, 9, 0, 0);
        dialog.AcceptButton = apply;
        dialog.Controls.Add(layout);
        PaintTheme(dialog);
        dialog.HandleCreated += (_, _) => SetTitleBar(dialog);
        dialog.ResumeLayout(true);
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;
        settings.Theme = Selected(theme) ?? "dark";
        settings.Language = Selected(languages) ?? "English";
        UpdateLabels();
        Render();
        PaintTheme(this);
        SaveAndRender();
    }

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
            _ when Equals(control.Tag, "signal") => palette.Signal,
            _ => palette.Ink
        };
        if (control is DesktopButton button)
            button.Palette = palette;
        if (control is DesktopToggle toggle)
            toggle.Palette = palette;
        if (control is Form form && form.IsHandleCreated)
            SetTitleBar(form);
        foreach (Control child in control.Controls)
            PaintTheme(child);
        control.Invalidate();
    }

    private void SetTitleBar(Form form)
    {
        int dark = settings.IsLight ? 0 : 1;
        _ = DwmSetWindowAttribute(form.Handle, 20, ref dark, sizeof(int));
    }

    private static bool IsAdministrator()
    {
        using var identity = WindowsIdentity.GetCurrent();
        return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static TableLayoutPanel Grid(int columns, int rows) => new()
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

    private static string? Selected(DesktopPicker picker) => picker.SelectedKey;
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
