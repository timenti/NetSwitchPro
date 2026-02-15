using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Text.Json;
using NetworkAdapterSwitcher.Models;
using NetworkAdapterSwitcher.Services;

namespace NetworkAdapterSwitcher;

internal sealed class MainForm : Form
{
    private const string SettingsFileName = "settings.json";
    private const string SettingsFolderName = "NetSwitchPro";

    private readonly NetshAdapterService _adapterService = new();

    private readonly Panel _selectionScreen = new() { Dock = DockStyle.Fill, Padding = new Padding(18, 16, 18, 14) };
    private readonly Panel _switchScreen = new() { Dock = DockStyle.Fill, Padding = new Padding(18, 16, 18, 14), Visible = false };

    private readonly Label _selectionTitle = new();
    private readonly Label _selectionSubtitle = new();
    private readonly Label _selectionListTitle = new();
    private readonly Label _selectionStatus = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true };

    private readonly Label _switchTitle = new();

    private readonly ThemedButton _updateButton = new("UPDATE", false);
    private readonly ThemedButton _themeButton = new("THEME", false);
    private readonly ThemedButton _okButton = new("OK", true);
    private readonly ThemedButton _settingsButton = new("⚙ Settings", false);
    private readonly ThemedButton _switchButton = new("↻", true, true);
    private readonly ThemedButton _languageButton = new("LANGUAGE", false);

    private readonly CheckBox _showVirtualCheck = new() { AutoSize = true };
    private readonly CheckBox _showBluetoothCheck = new() { AutoSize = true };

    private readonly SurfacePanel _selectionContainer = new() { Dock = DockStyle.Fill, Padding = new Padding(12) };
    private readonly Panel _listViewport = new() { Dock = DockStyle.Fill, BackColor = Color.Transparent, Padding = new Padding(0, 4, 0, 0) };
    private readonly Panel _listContent = new() { Dock = DockStyle.Top, Height = 10 };
    private readonly ThemedScrollBar _listScroll = new() { Dock = DockStyle.Right, Width = 12, Visible = false };

    private readonly AdapterStateCard _primaryCard = new();
    private readonly AdapterStateCard _secondaryCard = new();

    private readonly RichTextBox _logBox = new()
    {
        Dock = DockStyle.Fill,
        Multiline = true,
        ReadOnly = true,
        BorderStyle = BorderStyle.None,
        ScrollBars = RichTextBoxScrollBars.None,
        Font = new Font("Consolas", 9.8F)
    };
    private readonly ThemedScrollBar _logScroll = new() { Dock = DockStyle.Right, Width = 12, Visible = false };

    private readonly ContextMenuStrip _languageMenu = new();
    private readonly SurfacePanel _logPanel = new() { Dock = DockStyle.Bottom, Height = 150, Padding = new Padding(10) };
    private readonly System.Windows.Forms.Timer _switchCooldownTimer = new() { Interval = 200 };

    private readonly List<AdapterChoiceCard> _choiceCards = [];
    private readonly List<string> _selectedAdapterNames = [];
    private List<NetworkAdapterInfo> _adapters = [];
    private List<NetworkAdapterInfo> _allAdapters = [];
    private bool _suppressSettingsPersistence;

    private UiTheme _theme = UiTheme.Dark;
    private UiLanguage _language = UiLanguage.English;
    private DateTime _nextSwitchAvailableAt = DateTime.MinValue;
    private int _lastVirtualAdapterCount;
    private int _lastBluetoothAdapterCount;
    private int _lastFoundAdapterCount;
    private const int SwitchCooldownSeconds = 5;

    public MainForm()
    {
        Text = "NetSwitch Pro | Powered by Timenti®";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(860, 640);
        MinimumSize = new Size(860, 640);
        MaximumSize = new Size(860, 640);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        MinimizeBox = true;
        Font = new Font("Segoe UI", 10F);
        DoubleBuffered = true;
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
        UpdateStyles();
        ShowIcon = true;

        try
        {
            Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }
        catch
        {
            // keep default icon when extraction fails
        }

        BuildLanguageMenu();
        BuildSelectionScreen();
        BuildSwitchScreen();

        Controls.Add(_switchScreen);
        Controls.Add(_selectionScreen);

        _updateButton.Click += async (_, _) => await RefreshAdaptersAsync();
        _okButton.Click += (_, _) => OpenSwitchScreen();
        _settingsButton.Click += (_, _) => { ShowScreen(_switchScreen, _selectionScreen); AdjustWindowHeight(); };
        _themeButton.Click += (_, _) => ToggleTheme();
        _switchButton.Click += async (_, _) => await ExecuteSwitchAsync();
        _languageButton.Click += (_, _) =>
        {
            var x = Math.Max(0, (_languageButton.Width - _languageMenu.Width) / 2);
            _languageMenu.Show(_languageButton, new Point(x, _languageButton.Height - 1));
        };
        _showVirtualCheck.CheckedChanged += async (_, _) =>
        {
            SaveSettings();
            await RefreshAdaptersAsync();
        };
        _showBluetoothCheck.CheckedChanged += async (_, _) =>
        {
            SaveSettings();
            await RefreshAdaptersAsync();
        };

        _listScroll.ValueChanged += (_, value) => _listContent.Top = -value;
        _listViewport.MouseWheel += (_, e) =>
        {
            if (!_listScroll.Visible)
            {
                return;
            }

            var delta = e.Delta > 0 ? -32 : 32;
            _listScroll.SetValueClamped(_listScroll.Value + delta);
        };

        _logScroll.ValueChanged += (_, value) =>
        {
            NativeMethods.SetScrollPos(_logBox.Handle, NativeMethods.SbVert, value, true);
            NativeMethods.SendMessage(_logBox.Handle, NativeMethods.WmVscroll, (IntPtr)(NativeMethods.SbThumbPosition + 0x10000 * value), IntPtr.Zero);
        };

        _switchCooldownTimer.Tick += (_, _) =>
        {
            if (DateTime.UtcNow >= _nextSwitchAvailableAt)
            {
                _switchCooldownTimer.Stop();
                _switchButton.OverrideColor = null;
                _switchButton.Invalidate();
            }
        };

        Load += async (_, _) =>
        {
            LoadSettings();
            ApplyTheme();
            ApplyLanguage();
            await RefreshAdaptersAsync();

            if (_selectedAdapterNames.Count == 2)
            {
                OpenSwitchScreen();
            }
        };

        Resize += (_, _) => UpdateListLayout();
        FormClosing += (_, _) => SaveSettings();
    }

    protected override CreateParams CreateParams
    {
        get
        {
            const int WsExComposited = 0x02000000;
            var cp = base.CreateParams;
            cp.ExStyle |= WsExComposited;
            return cp;
        }
    }


    private void BuildSelectionScreen()
    {
        var header = new TableLayoutPanel { Dock = DockStyle.Top, Height = 96, ColumnCount = 2, BackColor = Color.Transparent };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));

        var headerTextPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
        _selectionTitle.Dock = DockStyle.Top;
        _selectionTitle.Height = 48;
        _selectionTitle.Font = new Font("Segoe UI", 24F, FontStyle.Bold);
        _selectionSubtitle.Dock = DockStyle.Top;
        _selectionSubtitle.Height = 30;
        _selectionSubtitle.Font = new Font("Segoe UI", 13F);
        headerTextPanel.Controls.Add(_selectionSubtitle);
        headerTextPanel.Controls.Add(_selectionTitle);

        var rightControls = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            BackColor = Color.Transparent,
            Padding = new Padding(0, 8, 0, 0)
        };

        _languageButton.Width = 140;
        _languageButton.Height = 38;
        _languageButton.Margin = new Padding(0, 0, 0, 0);

        _themeButton.Width = 140;
        _themeButton.Height = 38;
        _themeButton.Margin = new Padding(4, 0, 0, 0);

        rightControls.Controls.Add(_languageButton);
        rightControls.Controls.Add(_themeButton);

        header.Controls.Add(headerTextPanel, 0, 0);
        header.Controls.Add(rightControls, 1, 0);

        _selectionListTitle.Dock = DockStyle.Top;
        _selectionListTitle.Height = 28;
        _selectionListTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);

        _listViewport.Controls.Add(_listContent);
        _selectionContainer.Controls.Add(_listViewport);
        _selectionContainer.Controls.Add(_listScroll);
        _selectionContainer.Controls.Add(_selectionListTitle);

        var footer = new TableLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 96,
            ColumnCount = 3,
            BackColor = Color.Transparent,
            Padding = new Padding(0, 8, 0, 0)
        };
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 184));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 184));

        _updateButton.Dock = DockStyle.None;
        _updateButton.Size = new Size(168, 44);
        _updateButton.Anchor = AnchorStyles.None;

        _okButton.Dock = DockStyle.None;
        _okButton.Size = new Size(168, 44);
        _okButton.Anchor = AnchorStyles.None;
        _okButton.Enabled = false;

        var footerLeft = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, BackColor = Color.Transparent, Padding = new Padding(0, 2, 0, 0) };
        footerLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        footerLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 4));
        footerLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
        footerLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 2));
        footerLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
        footerLeft.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _showVirtualCheck.AutoSize = true;
        _showVirtualCheck.Margin = new Padding(0);
        _showVirtualCheck.Anchor = AnchorStyles.Left | AnchorStyles.Top;

        _showBluetoothCheck.AutoSize = true;
        _showBluetoothCheck.Margin = new Padding(0);
        _showBluetoothCheck.Anchor = AnchorStyles.Left | AnchorStyles.Top;

        _selectionStatus.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

        footerLeft.Controls.Add(_selectionStatus, 0, 0);
        footerLeft.Controls.Add(_showVirtualCheck, 0, 2);
        footerLeft.Controls.Add(_showBluetoothCheck, 0, 4);

        var updateWrap = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
        updateWrap.Controls.Add(_updateButton);
        updateWrap.Resize += (_, _) =>
        {
            _updateButton.Left = (updateWrap.Width - _updateButton.Width) / 2;
            _updateButton.Top = (updateWrap.Height - _updateButton.Height) / 2;
        };

        var okWrap = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
        okWrap.Controls.Add(_okButton);
        okWrap.Resize += (_, _) =>
        {
            _okButton.Left = (okWrap.Width - _okButton.Width) / 2;
            _okButton.Top = (okWrap.Height - _okButton.Height) / 2;
        };

        footer.Controls.Add(footerLeft, 0, 0);
        footer.Controls.Add(updateWrap, 1, 0);
        footer.Controls.Add(okWrap, 2, 0);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, BackColor = Color.Transparent };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 104));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 104));
        root.Controls.Add(header, 0, 0);
        root.Controls.Add(_selectionContainer, 0, 1);
        root.Controls.Add(footer, 0, 2);

        _selectionScreen.Controls.Add(root);
    }

    private void BuildSwitchScreen()
    {
        var header = new TableLayoutPanel { Dock = DockStyle.Top, Height = 52, ColumnCount = 2, BackColor = Color.Transparent };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 80));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));

        _switchTitle.Dock = DockStyle.Fill;
        _switchTitle.Font = new Font("Segoe UI", 22F, FontStyle.Bold);
        _switchTitle.TextAlign = ContentAlignment.MiddleLeft;
        _settingsButton.Dock = DockStyle.Fill;

        header.Controls.Add(_switchTitle, 0, 0);
        header.Controls.Add(_settingsButton, 1, 0);

        var centerPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
        _switchButton.Size = new Size(90, 90);
        _switchButton.Font = new Font("Segoe UI Symbol", 22F, FontStyle.Bold);
        centerPanel.Controls.Add(_switchButton);
        centerPanel.Resize += (_, _) =>
        {
            _switchButton.Left = (centerPanel.Width - _switchButton.Width) / 2;
            _switchButton.Top = (centerPanel.Height - _switchButton.Height) / 2;
        };

        var cards = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(0, 14, 0, 8), BackColor = Color.Transparent };
        cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43));
        cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14));
        cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43));

        _primaryCard.Dock = DockStyle.Fill;
        _secondaryCard.Dock = DockStyle.Fill;

        cards.Controls.Add(_primaryCard, 0, 0);
        cards.Controls.Add(centerPanel, 1, 0);
        cards.Controls.Add(_secondaryCard, 2, 0);

        _logPanel.Controls.Add(_logScroll);
        _logPanel.Controls.Add(_logBox);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, BackColor = Color.Transparent };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        root.Controls.Add(header, 0, 0);
        root.Controls.Add(cards, 0, 1);
        root.Controls.Add(_logPanel, 0, 2);

        _switchScreen.Controls.Add(root);
    }

    private async Task RefreshAdaptersAsync()
    {
        SetBusy(true, lockTopButtons: false);
        SetSelectionStatus(T(UiText.LoadingAdapters), false);

        try
        {
            var all = await Task.Run(() => _adapterService.GetAdapters().ToList());
            _allAdapters = all;
            _lastVirtualAdapterCount = all.Count(a => a.IsVirtual);
            _lastBluetoothAdapterCount = all.Count(a => a.IsBluetooth);
            _adapters = all.Where(a => (_showVirtualCheck.Checked || !a.IsVirtual) && (_showBluetoothCheck.Checked || !a.IsBluetooth)).ToList();
            _lastFoundAdapterCount = _adapters.Count;

            var selectedBeforeFilter = _selectedAdapterNames.Count;
            _selectedAdapterNames.RemoveAll(name => !_allAdapters.Any(a => string.Equals(a.Name, name, StringComparison.OrdinalIgnoreCase)));
            _selectedAdapterNames.RemoveAll(name => !_adapters.Any(a => string.Equals(a.Name, name, StringComparison.OrdinalIgnoreCase)));
            if (_selectedAdapterNames.Count != selectedBeforeFilter)
            {
                SaveSettings();
            }

            BuildAdapterCards();
            UpdateSwitchCards();
            UpdateSelectionStats(false);
        }
        catch (Exception ex)
        {
            SetSelectionStatus($"{T(UiText.Error)}: {ex.Message}", true);
            MessageBox.Show(this, ex.Message, T(UiText.Error), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false, lockTopButtons: false);
        }
    }

    private void BuildAdapterCards()
    {
        _listContent.SuspendLayout();
        _listContent.Controls.Clear();
        _choiceCards.Clear();

        var y = 0;
        foreach (var adapter in _adapters)
        {
            var card = new AdapterChoiceCard();
            card.SetTheme(_theme);
            card.SetAdapter(adapter);
            card.Width = _listViewport.Width - (_listScroll.Visible ? _listScroll.Width + 8 : 8);
            card.Height = 78;
            card.Left = 0;
            card.Top = y;
            y += card.Height + 8;

            card.IsSelected = _selectedAdapterNames.Contains(adapter.Name, StringComparer.OrdinalIgnoreCase);
            card.CardPressed += (_, _) => ToggleAdapter(adapter.Name);

            _choiceCards.Add(card);
            _listContent.Controls.Add(card);
        }

        _listContent.Height = Math.Max(1, y > 0 ? y - 8 : 1);
        _listContent.ResumeLayout();

        UpdateListLayout();
    }

    private void UpdateListLayout()
    {
        foreach (var card in _choiceCards)
        {
            card.Width = _listViewport.Width - (_listScroll.Visible ? _listScroll.Width + 8 : 8);
        }

        var viewportHeight = _listViewport.Height;
        var contentHeight = _listContent.Height;

        var needScroll = contentHeight > viewportHeight;
        _listScroll.Visible = needScroll;

        if (needScroll)
        {
            _listScroll.Maximum = Math.Max(0, contentHeight - viewportHeight);
            _listScroll.ViewportSize = Math.Max(1, viewportHeight);
            _listScroll.SetValueClamped(_listScroll.Value);
        }
        else
        {
            _listScroll.SetValueClamped(0);
            _listContent.Top = 0;
        }

        AdjustWindowHeight();
    }

    private void AdjustWindowHeight()
    {
        var desiredCards = Math.Max(_adapters.Count, 2);
        if (_adapters.Count > 6)
        {
            desiredCards = 6;
        }

        var cardAreaHeight = desiredCards * 86 + 82;
        var desired = 220 + cardAreaHeight + 76;
        desired = Math.Clamp(desired, 500, 760);

        var fixedSize = _switchScreen.Visible ? new Size(860, 640) : new Size(860, desired);
        if (ClientSize != fixedSize)
        {
            ClientSize = fixedSize;
            MinimumSize = fixedSize;
            MaximumSize = fixedSize;
        }
    }

    private void ToggleAdapter(string adapterName)
    {
        if (_selectedAdapterNames.Contains(adapterName, StringComparer.OrdinalIgnoreCase))
        {
            _selectedAdapterNames.RemoveAll(n => string.Equals(n, adapterName, StringComparison.OrdinalIgnoreCase));
        }
        else
        {
            if (_selectedAdapterNames.Count == 2)
            {
                _selectedAdapterNames.RemoveAt(0);
            }

            _selectedAdapterNames.Add(adapterName);
        }

        foreach (var card in _choiceCards)
        {
            card.IsSelected = _selectedAdapterNames.Contains(card.AdapterName, StringComparer.OrdinalIgnoreCase);
            card.Invalidate();
        }

        _okButton.Enabled = _selectedAdapterNames.Count == 2;
        UpdateSelectionStats(false);
        UpdateSwitchCards();
        SaveSettings();
    }

    private void OpenSwitchScreen()
    {
        if (_selectedAdapterNames.Count != 2)
        {
            MessageBox.Show(this, T(UiText.SelectExactlyTwo), T(UiText.Validation), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        ShowScreen(_selectionScreen, _switchScreen);
        AdjustWindowHeight();
        Log(T(UiText.InterfaceReady));
    }

    private async Task ExecuteSwitchAsync()
    {
        var cooldownLeft = _nextSwitchAvailableAt - DateTime.UtcNow;
        if (cooldownLeft > TimeSpan.Zero)
        {
            var seconds = Math.Max(1, (int)Math.Ceiling(cooldownLeft.TotalSeconds));
            Log(string.Format(T(UiText.NextSwitchAvailableIn), seconds), _theme.Error);
            return;
        }

        if (_selectedAdapterNames.Count != 2)
        {
            ShowScreen(_switchScreen, _selectionScreen);
            AdjustWindowHeight();
            return;
        }

        SetBusy(true);

        try
        {
            var firstName = _selectedAdapterNames[0];
            var secondName = _selectedAdapterNames[1];
            Log($"{T(UiText.Switching)}: '{firstName}' <-> '{secondName}'");

            var current = BuildAdapterLookup(await Task.Run(() => _adapterService.GetAdapters()));
            if (!current.TryGetValue(firstName, out var first) || !current.TryGetValue(secondName, out var second))
            {
                throw new InvalidOperationException(T(UiText.AdapterNotFound));
            }

            var target = DecideTargetState(first, second);
            Log($"{T(UiText.ApplyAction)}: {firstName}={(target.EnableFirst ? "Enable" : "Disable")}, {secondName}={(target.EnableSecond ? "Enable" : "Disable")}");

            await Task.Run(() =>
            {
                _adapterService.SetAdapterState(first.Name, target.EnableFirst);
                _adapterService.SetAdapterState(second.Name, target.EnableSecond);
            });

            await RefreshAdaptersAsync();
            Log(T(UiText.SwitchDone));
            AnimateSwitchIcon();

            _nextSwitchAvailableAt = DateTime.UtcNow.AddSeconds(SwitchCooldownSeconds);
            _switchButton.OverrideColor = Color.FromArgb(130, 130, 130);
            _switchButton.Invalidate();
            _switchCooldownTimer.Start();
            Log(string.Format(T(UiText.NextSwitchAvailableIn), SwitchCooldownSeconds), _theme.Error);
        }
        catch (Exception ex)
        {
            Log($"{T(UiText.Error)}: {ex.Message}");
            MessageBox.Show(this, ex.Message, T(UiText.Error), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void UpdateSwitchCards()
    {
        var map = BuildAdapterLookup(_allAdapters.Count > 0 ? _allAdapters : _adapters);

        if (_selectedAdapterNames.Count > 0 && map.TryGetValue(_selectedAdapterNames[0], out var first))
        {
            _primaryCard.SetAdapter(first, true);
        }
        else
        {
            _primaryCard.SetPlaceholder(T(UiText.SelectPrimary), T(UiText.OpenSettings));
        }

        if (_selectedAdapterNames.Count > 1 && map.TryGetValue(_selectedAdapterNames[1], out var second))
        {
            _secondaryCard.SetAdapter(second, false);
        }
        else
        {
            _secondaryCard.SetPlaceholder(T(UiText.SelectSecondary), T(UiText.OpenSettings));
        }
    }


    private static Dictionary<string, NetworkAdapterInfo> BuildAdapterLookup(IEnumerable<NetworkAdapterInfo> adapters)
    {
        var map = new Dictionary<string, NetworkAdapterInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var adapter in adapters)
        {
            if (!map.ContainsKey(adapter.Name))
            {
                map[adapter.Name] = adapter;
            }
        }

        return map;
    }

    private static (bool EnableFirst, bool EnableSecond) DecideTargetState(NetworkAdapterInfo first, NetworkAdapterInfo second)
    {
        if (first.IsAdminEnabled && !second.IsAdminEnabled)
        {
            return (false, true);
        }

        if (!first.IsAdminEnabled && second.IsAdminEnabled)
        {
            return (true, false);
        }

        if (first.IsAdminEnabled && second.IsAdminEnabled)
        {
            return (false, true);
        }

        return (true, false);
    }

    private void ShowScreen(Control from, Control to)
    {
        from.Visible = false;
        to.Visible = true;
        to.BringToFront();
    }

    private void ToggleTheme()
    {
        _theme = _theme.Name == "dark" ? UiTheme.Light : UiTheme.Dark;
        SaveSettings();
        ApplyTheme();
    }

    private void ApplyTheme()
    {
        BackColor = _theme.BgA;
        _selectionScreen.BackColor = _theme.BgA;
        _switchScreen.BackColor = _theme.BgA;

        _updateButton.SetTheme(_theme);
        _themeButton.SetTheme(_theme);
        _okButton.SetTheme(_theme);
        _settingsButton.SetTheme(_theme);
        _switchButton.SetTheme(_theme);
        _languageButton.SetTheme(_theme);
        _listScroll.SetTheme(_theme);

        _selectionContainer.SetTheme(_theme);
        _logPanel.SetTheme(_theme);
        var listBackground = _theme.Name == "light" ? Color.FromArgb(242, 246, 251) : _theme.Surface;
        _listViewport.BackColor = listBackground;
        _listContent.BackColor = listBackground;
        _primaryCard.SetTheme(_theme);
        _secondaryCard.SetTheme(_theme);

        foreach (var card in _choiceCards)
        {
            card.SetTheme(_theme);
        }

        _selectionStatus.ForeColor = _theme.Muted;
        _selectionTitle.ForeColor = _theme.Text;
        _selectionSubtitle.ForeColor = _theme.Muted;
        _selectionListTitle.ForeColor = _theme.Text;
        _switchTitle.ForeColor = _theme.Text;

        _showVirtualCheck.ForeColor = _theme.Muted;
        _showVirtualCheck.BackColor = Color.Transparent;
        _showBluetoothCheck.ForeColor = _theme.Muted;
        _showBluetoothCheck.BackColor = Color.Transparent;

        ApplyDarkTitleBar();
        Text = "NetSwitch Pro | Powered by Timenti®";

        _themeButton.Text = T(UiText.Theme).ToUpperInvariant();

        _languageMenu.BackColor = _theme.Surface;
        _languageMenu.ForeColor = _theme.Text;
        _languageMenu.Renderer = new ThemeMenuRenderer(_theme);
        _languageMenu.Width = Math.Max(110, _languageButton.Width - 18);
        foreach (ToolStripItem item in _languageMenu.Items)
        {
            item.BackColor = _theme.Surface;
            item.ForeColor = _theme.Text;
        }

        _logBox.BackColor = _theme.Surface;
        _logBox.ForeColor = _theme.Muted;
        _logScroll.SetTheme(_theme);
        SyncLogScrollbar();
    }

    private void BuildLanguageMenu()
    {
        _languageMenu.ShowImageMargin = false;
        _languageMenu.ShowCheckMargin = false;
        _languageMenu.AutoSize = false;
        _languageMenu.Padding = new Padding(0, 2, 0, 6);
        _languageMenu.DropShadowEnabled = true;
        _languageMenu.Opening += (_, _) =>
        {
            var width = Math.Max(110, _languageButton.Width - 18);
            _languageMenu.Width = width;
            foreach (ToolStripItem menuItem in _languageMenu.Items)
            {
                menuItem.AutoSize = false;
                menuItem.Width = width - 2;
                menuItem.Height = 28;
            }

            _languageMenu.Height = _languageMenu.Items.Count * 28 + _languageMenu.Padding.Vertical + 6;

        };

        AddLanguageMenuItem(UiLanguage.English, "English");
        AddLanguageMenuItem(UiLanguage.Russian, "Russian");
        AddLanguageMenuItem(UiLanguage.French, "French");
        AddLanguageMenuItem(UiLanguage.German, "German");
        AddLanguageMenuItem(UiLanguage.Chinese, "Chinese");
        AddLanguageMenuItem(UiLanguage.Polish, "Polish");
        AddLanguageMenuItem(UiLanguage.Japanese, "Japanese");
    }

    private void AddLanguageMenuItem(UiLanguage lang, string name)
    {
        var item = new ToolStripMenuItem(name)
        {
            ImageScaling = ToolStripItemImageScaling.None
        };
        item.Click += (_, _) =>
        {
            _language = lang;
            SaveSettings();
            ApplyLanguage();
        };
        _languageMenu.Items.Add(item);
    }

    private void ApplyLanguage()
    {
        _selectionTitle.Text = T(UiText.SystemConfiguration);
        _selectionSubtitle.Text = T(UiText.SelectExactlyTwo);
        _selectionListTitle.Text = T(UiText.AvailableAdapters);
        _switchTitle.Text = T(UiText.SwitchController);

        UpdateAdapterFilterCaptions();
        _languageButton.Text = "LANGUAGE";

        _updateButton.Text = T(UiText.Update).ToUpperInvariant();
        _okButton.Text = T(UiText.Ok).ToUpperInvariant();
        _settingsButton.Text = "⚙ " + T(UiText.Settings).ToUpperInvariant();

        _themeButton.Text = T(UiText.Theme).ToUpperInvariant();

        _primaryCard.SetKind(T(UiText.PrimaryInterface));
        _secondaryCard.SetKind(T(UiText.SecondaryInterface));

        UpdateSelectionStats(false);
    }

    private void UpdateAdapterFilterCaptions()
    {
        _showVirtualCheck.Text = $"{T(UiText.ShowVirtualAdapters)}: {_lastVirtualAdapterCount}";
        _showBluetoothCheck.Text = $"{T(UiText.ShowBluetoothAdapters)}: {_lastBluetoothAdapterCount}";
    }

    private void UpdateSelectionStats(bool isError)
    {
        SetSelectionStatus($"{T(UiText.FoundAdapters)}: {_lastFoundAdapterCount}. {T(UiText.Selected)}: {_selectedAdapterNames.Count}/2", isError);
        UpdateAdapterFilterCaptions();
    }

    private string T(UiText key) => Localizer.Get(_language, key);

    private void SetBusy(bool busy, bool lockTopButtons = true)
    {
        _updateButton.Enabled = lockTopButtons ? !busy : true;
        _themeButton.Enabled = lockTopButtons ? !busy : true;
        _okButton.Enabled = lockTopButtons ? (!busy && _selectedAdapterNames.Count == 2) : _selectedAdapterNames.Count == 2;
        _settingsButton.Enabled = !busy;
        _switchButton.Enabled = !busy;
        _languageButton.Enabled = lockTopButtons ? !busy : true;
        _showVirtualCheck.Enabled = !busy;
        _showBluetoothCheck.Enabled = !busy;

        foreach (var card in _choiceCards)
        {
            card.Enabled = !busy;
        }
    }

    private void SetSelectionStatus(string text, bool isError)
    {
        _selectionStatus.Text = text;
        _selectionStatus.ForeColor = isError ? _theme.Error : _theme.Muted;
    }

    private void Log(string line, Color? color = null)
    {
        _logBox.SelectionStart = _logBox.TextLength;
        _logBox.SelectionLength = 0;
        _logBox.SelectionColor = color ?? _theme.Muted;
        _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {line}{Environment.NewLine}");
        _logBox.SelectionColor = _theme.Muted;
        _logBox.SelectionStart = _logBox.TextLength;
        _logBox.ScrollToCaret();
        SyncLogScrollbar();
    }

    private void SyncLogScrollbar()
    {
        if (!_logBox.IsHandleCreated)
        {
            return;
        }

        var info = NativeMethods.GetScrollInfo(_logBox.Handle);
        var max = Math.Max(0, info.nMax - (int)info.nPage);
        _logScroll.Maximum = max;
        _logScroll.ViewportSize = Math.Max(1, (int)info.nPage);
        _logScroll.Visible = info.nMax > (int)info.nPage && max > 0;
        _logScroll.SetValueClamped(Math.Clamp(info.nPos, 0, max));
    }

    private void AnimateSwitchIcon()
    {
        var frames = new[] { "↻", "↺", "↻", "↺", "↻" };
        var idx = 0;
        var timer = new System.Windows.Forms.Timer { Interval = 85 };
        timer.Tick += (_, _) =>
        {
            _switchButton.Text = frames[idx++];
            if (idx >= frames.Length)
            {
                timer.Stop();
                timer.Dispose();
                _switchButton.Text = "↻";
            }
        };
        timer.Start();
    }

    private string PrimarySettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        SettingsFolderName,
        SettingsFileName);

    private string SecondarySettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Temp",
        SettingsFolderName,
        SettingsFileName);

    private string FallbackSettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        SettingsFolderName,
        SettingsFileName);

    private IEnumerable<string> CandidateSettingsPaths
    {
        get
        {
            yield return PrimarySettingsPath;
            yield return SecondarySettingsPath;
            yield return FallbackSettingsPath;
        }
    }

    private void SaveSettings()
    {
        if (_suppressSettingsPersistence)
        {
            return;
        }

        var settings = new AppSettings
        {
            Theme = _theme.Name,
            Language = _language.ToString(),
            ShowVirtual = _showVirtualCheck.Checked,
            ShowBluetooth = _showBluetoothCheck.Checked,
            AdapterA = _selectedAdapterNames.ElementAtOrDefault(0),
            AdapterB = _selectedAdapterNames.ElementAtOrDefault(1)
        };

        TryWriteSettings(settings);
    }

    private void TryWriteSettings(AppSettings settings)
    {
        var serialized = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });

        foreach (var path in CandidateSettingsPaths)
        {
            _ = TryWriteSettingsToPath(path, serialized);
        }
    }

    private bool TryWriteSettingsToPath(string targetPath, string serializedSettings)
    {
        try
        {
            var settingsDirectory = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(settingsDirectory))
            {
                Directory.CreateDirectory(settingsDirectory);
            }

            if (File.Exists(targetPath))
            {
                var existingAttributes = File.GetAttributes(targetPath);
                if ((existingAttributes & FileAttributes.ReadOnly) != 0)
                {
                    File.SetAttributes(targetPath, existingAttributes & ~FileAttributes.ReadOnly);
                }
            }

            var tempPath = targetPath + ".tmp";
            File.WriteAllText(tempPath, serializedSettings);

            if (File.Exists(targetPath))
            {
                File.Delete(targetPath);
            }

            File.Move(tempPath, targetPath);

            try
            {
                File.SetAttributes(targetPath, File.GetAttributes(targetPath) | FileAttributes.Hidden);
            }
            catch
            {
                // ignore if hidden attribute cannot be set
            }

            return true;
        }
        catch
        {
            return false;
        }
    }


    private IEnumerable<string> GetExistingSettingsPathsOrdered()
    {
        var existing = new List<(string Path, DateTime WriteTime)>();

        foreach (var path in CandidateSettingsPaths)
        {
            try
            {
                if (File.Exists(path))
                {
                    existing.Add((path, File.GetLastWriteTimeUtc(path)));
                }
            }
            catch
            {
                // ignore unreadable candidate and continue
            }
        }

        return existing
            .OrderByDescending(x => x.WriteTime)
            .Select(x => x.Path);
    }

    private void LoadSettings()
    {
        var candidatePaths = GetExistingSettingsPathsOrdered().ToList();
        if (candidatePaths.Count == 0)
        {
            _language = UiLanguage.English;
            return;
        }

        AppSettings? settings = null;

        foreach (var settingsPath in candidatePaths)
        {
            try
            {
                settings = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(settingsPath));
                if (settings is not null)
                {
                    break;
                }
            }
            catch
            {
                // broken settings file, try older candidate
            }
        }

        if (settings is null)
        {
            _language = UiLanguage.English;
            return;
        }

        try
        {
            _suppressSettingsPersistence = true;

            _theme = settings.Theme?.ToLowerInvariant() == "light" ? UiTheme.Light : UiTheme.Dark;

            if (Enum.TryParse<UiLanguage>(settings.Language, true, out var lang))
            {
                _language = lang;
            }
            else
            {
                _language = UiLanguage.English;
            }

            _showVirtualCheck.Checked = settings.ShowVirtual;
            _showBluetoothCheck.Checked = settings.ShowBluetooth;

            _selectedAdapterNames.Clear();
            if (!string.IsNullOrWhiteSpace(settings.AdapterA))
            {
                _selectedAdapterNames.Add(settings.AdapterA);
            }
            if (!string.IsNullOrWhiteSpace(settings.AdapterB) && !string.Equals(settings.AdapterA, settings.AdapterB, StringComparison.OrdinalIgnoreCase))
            {
                _selectedAdapterNames.Add(settings.AdapterB);
            }
        }
        catch
        {
            // ignore broken settings
        }
        finally
        {
            _suppressSettingsPersistence = false;
        }
    }

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    private void ApplyDarkTitleBar()
    {
        try
        {
            var dark = _theme.Name == "dark" ? 1 : 0;
            const int DwmwaUseImmersiveDarkMode = 20;
            _ = DwmSetWindowAttribute(Handle, DwmwaUseImmersiveDarkMode, ref dark, sizeof(int));
        }
        catch
        {
            // best-effort
        }
    }

    private sealed class AppSettings
    {
        public string? Theme { get; set; }
        public string? Language { get; set; }
        public bool ShowVirtual { get; set; }
        public bool ShowBluetooth { get; set; }
        public string? AdapterA { get; set; }
        public string? AdapterB { get; set; }
    }
}

internal enum UiLanguage
{
    English,
    Russian,
    French,
    German,
    Chinese,
    Polish,
    Japanese
}

internal enum UiText
{
    SystemConfiguration,
    SelectExactlyTwo,
    AvailableAdapters,
    Update,
    Theme,
    Ok,
    Settings,
    Language,
    ShowVirtualAdapters,
    ShowBluetoothAdapters,
    FoundAdapters,
    Selected,
    Validation,
    Error,
    LoadingAdapters,
    InterfaceReady,
    Switching,
    ApplyAction,
    SwitchDone,
    AdapterNotFound,
    SwitchController,
    PrimaryInterface,
    SecondaryInterface,
    SelectPrimary,
    SelectSecondary,
    OpenSettings,
    NextSwitchAvailableIn
}

internal static class Localizer
{
    private static readonly Dictionary<UiLanguage, Dictionary<UiText, string>> Strings = new()
    {
        [UiLanguage.English] = new()
        {
            [UiText.SystemConfiguration] = "System Configuration",
            [UiText.SelectExactlyTwo] = "Select exactly 2 adapters to switch between.",
            [UiText.AvailableAdapters] = "Available adapters",
            [UiText.Update] = "Update",
            [UiText.Theme] = "Theme",
            [UiText.Ok] = "OK",
            [UiText.Settings] = "Settings",
            [UiText.Language] = "Language",
            [UiText.ShowVirtualAdapters] = "Virtual adapters",
            [UiText.ShowBluetoothAdapters] = "Bluetooth adapters",
            [UiText.FoundAdapters] = "Found adapters",
            [UiText.Selected] = "Selected",
            [UiText.Validation] = "Validation",
            [UiText.Error] = "Error",
            [UiText.LoadingAdapters] = "Scanning adapters...",
            [UiText.InterfaceReady] = "Interface is ready.",
            [UiText.Switching] = "Switching",
            [UiText.ApplyAction] = "Applying action",
            [UiText.SwitchDone] = "Switch completed successfully.",
            [UiText.AdapterNotFound] = "One of selected adapters was not found.",
            [UiText.SwitchController] = "NetSwitch Controller",
            [UiText.PrimaryInterface] = "PRIMARY INTERFACE",
            [UiText.SecondaryInterface] = "SECONDARY INTERFACE",
            [UiText.SelectPrimary] = "Select primary adapter",
            [UiText.SelectSecondary] = "Select secondary adapter",
            [UiText.OpenSettings] = "Open settings to choose adapters.",
            [UiText.NextSwitchAvailableIn] = "Next switch will be available in {0} sec."
        },
        [UiLanguage.Russian] = new()
        {
            [UiText.SystemConfiguration] = "Системная конфигурация",
            [UiText.SelectExactlyTwo] = "Выберите ровно 2 адаптера для переключения.",
            [UiText.AvailableAdapters] = "Доступные адаптеры",
            [UiText.Update] = "Обновить",
            [UiText.Theme] = "Тема",
            [UiText.Ok] = "ОК",
            [UiText.Settings] = "Настройки",
            [UiText.Language] = "Язык",
            [UiText.ShowVirtualAdapters] = "Виртуальные адаптеры",
            [UiText.ShowBluetoothAdapters] = "Bluetooth адаптеры",
            [UiText.FoundAdapters] = "Найдено адаптеров",
            [UiText.Selected] = "Выбрано",
            [UiText.Validation] = "Проверка",
            [UiText.Error] = "Ошибка",
            [UiText.LoadingAdapters] = "Сканирование адаптеров...",
            [UiText.InterfaceReady] = "Интерфейс готов.",
            [UiText.Switching] = "Переключение",
            [UiText.ApplyAction] = "Применение",
            [UiText.SwitchDone] = "Переключение выполнено успешно.",
            [UiText.AdapterNotFound] = "Один из выбранных адаптеров не найден.",
            [UiText.SwitchController] = "Контроллер NetSwitch",
            [UiText.PrimaryInterface] = "ОСНОВНОЙ АДАПТЕР",
            [UiText.SecondaryInterface] = "ВТОРИЧНЫЙ АДАПТЕР",
            [UiText.SelectPrimary] = "Выберите основной адаптер",
            [UiText.SelectSecondary] = "Выберите вторичный адаптер",
            [UiText.OpenSettings] = "Откройте настройки для выбора адаптеров.",
            [UiText.NextSwitchAvailableIn] = "Следующее переключение будет доступно через {0} секунд."
        },
        [UiLanguage.French] = new()
        {
            [UiText.SystemConfiguration] = "Configuration du système",
            [UiText.SelectExactlyTwo] = "Sélectionnez exactement 2 adaptateurs.",
            [UiText.AvailableAdapters] = "Adaptateurs disponibles",
            [UiText.Update] = "Mettre à jour",
            [UiText.Theme] = "Thème",
            [UiText.Ok] = "OK",
            [UiText.Settings] = "Paramètres",
            [UiText.Language] = "Langue",
            [UiText.ShowVirtualAdapters] = "Virtuel",
            [UiText.ShowBluetoothAdapters] = "Bluetooth",
            [UiText.FoundAdapters] = "Adaptateurs trouvés",
            [UiText.Selected] = "Sélectionnés",
            [UiText.Validation] = "Validation",
            [UiText.Error] = "Erreur",
            [UiText.LoadingAdapters] = "Analyse des adaptateurs...",
            [UiText.InterfaceReady] = "Interface prête.",
            [UiText.Switching] = "Commutation",
            [UiText.ApplyAction] = "Application",
            [UiText.SwitchDone] = "Commutation terminée.",
            [UiText.AdapterNotFound] = "Adaptateur introuvable.",
            [UiText.SwitchController] = "Contrôleur NetSwitch",
            [UiText.PrimaryInterface] = "ADAPTATEUR PRINCIPAL",
            [UiText.SecondaryInterface] = "ADAPTATEUR SECONDAIRE",
            [UiText.SelectPrimary] = "Sélectionnez l'adaptateur principal",
            [UiText.SelectSecondary] = "Sélectionnez l'adaptateur secondaire",
            [UiText.OpenSettings] = "Ouvrez les paramètres pour choisir."
        },
        [UiLanguage.German] = new()
        {
            [UiText.SystemConfiguration] = "Systemkonfiguration",
            [UiText.SelectExactlyTwo] = "Wählen Sie genau 2 Adapter aus.",
            [UiText.AvailableAdapters] = "Verfügbare Adapter",
            [UiText.Update] = "Aktualisieren",
            [UiText.Theme] = "Thema",
            [UiText.Ok] = "OK",
            [UiText.Settings] = "Einstellungen",
            [UiText.Language] = "Sprache",
            [UiText.ShowVirtualAdapters] = "Virtuell",
            [UiText.ShowBluetoothAdapters] = "Bluetooth",
            [UiText.FoundAdapters] = "Gefundene Adapter",
            [UiText.Selected] = "Ausgewählt",
            [UiText.Validation] = "Prüfung",
            [UiText.Error] = "Fehler",
            [UiText.LoadingAdapters] = "Adapter werden gesucht...",
            [UiText.InterfaceReady] = "Oberfläche ist bereit.",
            [UiText.Switching] = "Umschalten",
            [UiText.ApplyAction] = "Aktion anwenden",
            [UiText.SwitchDone] = "Umschalten abgeschlossen.",
            [UiText.AdapterNotFound] = "Adapter nicht gefunden.",
            [UiText.SwitchController] = "NetSwitch-Steuerung",
            [UiText.PrimaryInterface] = "PRIMÄRER ADAPTER",
            [UiText.SecondaryInterface] = "SEKUNDÄRER ADAPTER",
            [UiText.SelectPrimary] = "Primären Adapter wählen",
            [UiText.SelectSecondary] = "Sekundären Adapter wählen",
            [UiText.OpenSettings] = "Öffnen Sie Einstellungen zum Auswählen."
        },
        [UiLanguage.Chinese] = new()
        {
            [UiText.SystemConfiguration] = "系统配置",
            [UiText.SelectExactlyTwo] = "请选择正好 2 个适配器。",
            [UiText.AvailableAdapters] = "可用适配器",
            [UiText.Update] = "更新",
            [UiText.Theme] = "主题",
            [UiText.Ok] = "确定",
            [UiText.Settings] = "设置",
            [UiText.Language] = "语言",
            [UiText.ShowVirtualAdapters] = "虚拟",
            [UiText.ShowBluetoothAdapters] = "Bluetooth",
            [UiText.FoundAdapters] = "已找到适配器",
            [UiText.Selected] = "已选择",
            [UiText.Validation] = "验证",
            [UiText.Error] = "错误",
            [UiText.LoadingAdapters] = "正在扫描适配器...",
            [UiText.InterfaceReady] = "界面已就绪。",
            [UiText.Switching] = "切换",
            [UiText.ApplyAction] = "应用操作",
            [UiText.SwitchDone] = "切换成功完成。",
            [UiText.AdapterNotFound] = "未找到所选适配器。",
            [UiText.SwitchController] = "NetSwitch 控制器",
            [UiText.PrimaryInterface] = "主适配器",
            [UiText.SecondaryInterface] = "次适配器",
            [UiText.SelectPrimary] = "选择主适配器",
            [UiText.SelectSecondary] = "选择次适配器",
            [UiText.OpenSettings] = "打开设置以选择适配器。"
        },
        [UiLanguage.Polish] = new()
        {
            [UiText.SystemConfiguration] = "Konfiguracja systemu",
            [UiText.SelectExactlyTwo] = "Wybierz dokładnie 2 adaptery.",
            [UiText.AvailableAdapters] = "Dostępne adaptery",
            [UiText.Update] = "Aktualizuj",
            [UiText.Theme] = "Motyw",
            [UiText.Ok] = "OK",
            [UiText.Settings] = "Ustawienia",
            [UiText.Language] = "Język",
            [UiText.ShowVirtualAdapters] = "Wirtualne",
            [UiText.ShowBluetoothAdapters] = "Bluetooth",
            [UiText.FoundAdapters] = "Znalezione adaptery",
            [UiText.Selected] = "Wybrane",
            [UiText.Validation] = "Walidacja",
            [UiText.Error] = "Błąd",
            [UiText.LoadingAdapters] = "Skanowanie adapterów...",
            [UiText.InterfaceReady] = "Interfejs gotowy.",
            [UiText.Switching] = "Przełączanie",
            [UiText.ApplyAction] = "Zastosuj akcję",
            [UiText.SwitchDone] = "Przełączanie zakończone.",
            [UiText.AdapterNotFound] = "Nie znaleziono adaptera.",
            [UiText.SwitchController] = "Kontroler NetSwitch",
            [UiText.PrimaryInterface] = "GŁÓWNY ADAPTER",
            [UiText.SecondaryInterface] = "DRUGI ADAPTER",
            [UiText.SelectPrimary] = "Wybierz główny adapter",
            [UiText.SelectSecondary] = "Wybierz drugi adapter",
            [UiText.OpenSettings] = "Otwórz ustawienia, aby wybrać adaptery."
        },
        [UiLanguage.Japanese] = new()
        {
            [UiText.SystemConfiguration] = "システム設定",
            [UiText.SelectExactlyTwo] = "切替対象のアダプターを2つ選択してください。",
            [UiText.AvailableAdapters] = "利用可能なアダプター",
            [UiText.Update] = "更新",
            [UiText.Theme] = "テーマ",
            [UiText.Ok] = "OK",
            [UiText.Settings] = "設定",
            [UiText.Language] = "言語",
            [UiText.ShowVirtualAdapters] = "仮想",
            [UiText.ShowBluetoothAdapters] = "Bluetooth",
            [UiText.FoundAdapters] = "検出されたアダプター",
            [UiText.Selected] = "選択済み",
            [UiText.Validation] = "検証",
            [UiText.Error] = "エラー",
            [UiText.LoadingAdapters] = "アダプターをスキャン中...",
            [UiText.InterfaceReady] = "インターフェースの準備完了。",
            [UiText.Switching] = "切り替え",
            [UiText.ApplyAction] = "操作を適用",
            [UiText.SwitchDone] = "切り替えが完了しました。",
            [UiText.AdapterNotFound] = "選択されたアダプターが見つかりません。",
            [UiText.SwitchController] = "NetSwitch コントローラー",
            [UiText.PrimaryInterface] = "プライマリアダプター",
            [UiText.SecondaryInterface] = "セカンダリアダプター",
            [UiText.SelectPrimary] = "プライマリアダプターを選択",
            [UiText.SelectSecondary] = "セカンダリアダプターを選択",
            [UiText.OpenSettings] = "設定を開いてアダプターを選択してください。"
        }
    };

    public static string Get(UiLanguage lang, UiText key) =>
        Strings.TryGetValue(lang, out var dict) && dict.TryGetValue(key, out var value)
            ? value
            : Strings[UiLanguage.English][key];
}

internal readonly record struct UiTheme(string Name, string DisplayName, Color BgA, Color BgB, Color Surface, Color Stroke, Color Text, Color Muted, Color Accent, Color Error)
{
    public static UiTheme Dark => new(
        "dark",
        "Dark",
        Color.FromArgb(12, 15, 28),
        Color.FromArgb(14, 33, 46),
        Color.FromArgb(24, 29, 43),
        Color.FromArgb(80, 108, 144),
        Color.FromArgb(242, 246, 255),
        Color.FromArgb(156, 175, 203),
        Color.FromArgb(23, 223, 206),
        Color.FromArgb(244, 90, 104));

    public static UiTheme Light => new(
        "light",
        "Light",
        Color.FromArgb(236, 240, 247),
        Color.FromArgb(220, 227, 238),
        Color.FromArgb(248, 250, 253),
        Color.FromArgb(154, 170, 198),
        Color.FromArgb(30, 43, 66),
        Color.FromArgb(95, 112, 140),
        Color.FromArgb(0, 184, 255),
        Color.FromArgb(226, 52, 69));
}

internal class SurfacePanel : Panel
{
    public Color FillColor { get; protected set; } = UiTheme.Dark.Surface;
    public Color BorderColor { get; protected set; } = UiTheme.Dark.Stroke;
    protected UiTheme CurrentTheme { get; private set; } = UiTheme.Dark;
    protected virtual float BorderThickness => CurrentTheme.Name == "light" ? 4f : 2f;

    public SurfacePanel()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
    }

    public virtual void SetTheme(UiTheme theme)
    {
        CurrentTheme = theme;
        FillColor = theme.Surface;
        BorderColor = theme.Stroke;
        BackColor = theme.BgA;
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(Parent?.BackColor ?? BackColor);
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        var inset = (int)Math.Ceiling(BorderThickness);
        var rect = new Rectangle(inset, inset, Math.Max(1, Width - inset * 2 - 1), Math.Max(1, Height - inset * 2 - 1));
        using var path = RoundRect(rect, 14);
        using var fill = new SolidBrush(FillColor);
        using var pen = new Pen(BorderColor, BorderThickness) { Alignment = PenAlignment.Inset };

        e.Graphics.FillPath(fill, path);
        e.Graphics.DrawPath(pen, path);
    }

    protected override void OnResize(EventArgs eventargs)
    {
        base.OnResize(eventargs);
        var inset = (int)Math.Ceiling(BorderThickness);
        var rect = new Rectangle(inset, inset, Math.Max(1, Width - inset * 2 - 1), Math.Max(1, Height - inset * 2 - 1));
        using var path = RoundRect(rect, 14);
        Region = new Region(path);
        Invalidate();
    }

    internal static GraphicsPath RoundRect(Rectangle rect, int radius)
    {
        var d = Math.Max(2, radius * 2);
        var path = new GraphicsPath();
        path.AddArc(rect.X, rect.Y, d, d, 180, 90);
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }
}

internal sealed class ThemedButton : Button
{
    private readonly bool _filled;
    private readonly bool _round;
    private UiTheme _theme = UiTheme.Dark;
    private bool _hover;

    public Color? OverrideColor { get; set; }

    public ThemedButton(string text, bool filled, bool round = false)
    {
        Text = text;
        _filled = filled;
        _round = round;
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        Font = new Font("Segoe UI", 10.5F, FontStyle.Bold);
        Cursor = Cursors.Hand;
        Height = 40;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

        MouseEnter += (_, _) => { _hover = true; Invalidate(); };
        MouseLeave += (_, _) => { _hover = false; Invalidate(); };
    }

    public void SetTheme(UiTheme theme)
    {
        _theme = theme;
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(Parent?.BackColor ?? BackColor);
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        var rect = new Rectangle(1, 1, Width - 3, Height - 3);
        var radius = _round ? rect.Height / 2 : 10;

        var baseFill = _filled
            ? _theme.Accent
            : (_theme.Name == "light" ? Color.FromArgb(229, 233, 240) : Color.FromArgb(34, 40, 56));
        var fill = OverrideColor ?? baseFill;
        if (_hover)
        {
            fill = ControlPaint.Light(fill, 0.08f);
        }
        if (!Enabled)
        {
            fill = Color.FromArgb(120, 120, 120);
        }

        using var path = SurfacePanel.RoundRect(rect, radius);
        using var b = new SolidBrush(fill);
        var borderThickness = _theme.Name == "light" ? 4f : 2f;
        using var p = new Pen(Color.FromArgb(185, _theme.Stroke), borderThickness);

        e.Graphics.FillPath(b, path);
        e.Graphics.DrawPath(p, path);

        var textColor = _filled ? Color.FromArgb(8, 28, 34) : _theme.Text;
        if (!Enabled)
        {
            textColor = Color.FromArgb(145, textColor);
        }

        TextRenderer.DrawText(e.Graphics, Text, Font, rect, textColor,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    protected override void OnPaintBackground(PaintEventArgs pevent)
    {
        pevent.Graphics.Clear(Parent?.BackColor ?? BackColor);
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        var rect = new Rectangle(1, 1, Math.Max(1, Width - 3), Math.Max(1, Height - 3));
        var radius = _round ? Math.Max(1, rect.Height / 2) : 10;
        using var path = SurfacePanel.RoundRect(rect, radius);
        Region = new Region(path);
        Invalidate();
    }
}

internal sealed class AdapterChoiceCard : SurfacePanel
{
    private readonly Label _name = new() { AutoSize = true, Font = new Font("Segoe UI", 12.5F, FontStyle.Bold) };
    private readonly Label _meta = new() { AutoSize = true, Font = new Font("Segoe UI", 10F) };
    private readonly Label _state = new() { AutoSize = true, Font = new Font("Segoe UI", 11F, FontStyle.Bold) };

    private UiTheme _theme = UiTheme.Dark;

    public string AdapterName { get; private set; } = string.Empty;
    public bool IsSelected { get; set; }

    protected override float BorderThickness => IsSelected
        ? (CurrentTheme.Name == "light" ? 6f : 4f)
        : (CurrentTheme.Name == "light" ? 3f : 2f);

    public event EventHandler? CardPressed;
    private DateTime _lastPress = DateTime.MinValue;

    public AdapterChoiceCard()
    {
        Padding = new Padding(12, 8, 12, 6);
        Cursor = Cursors.Hand;

        var stack = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.Transparent };
        stack.Controls.Add(_name);
        stack.Controls.Add(_meta);
        stack.Controls.Add(_state);
        Controls.Add(stack);

        AttachPress(this);
        AttachPress(stack);
        foreach (Control child in stack.Controls)
        {
            AttachPress(child);
        }
    }

    public override void SetTheme(UiTheme theme)
    {
        _theme = theme;
        base.SetTheme(theme);
        _name.ForeColor = theme.Text;
        _meta.ForeColor = theme.Muted;
        if (_state.Text != "● CONNECTED" && _state.Text != "● DISABLED")
        {
            _state.ForeColor = theme.Muted;
        }
    }

    public void SetAdapter(NetworkAdapterInfo adapter)
    {
        AdapterName = adapter.Name;
        _name.Text = adapter.Name;
        _meta.Text = $"{adapter.Type} · {adapter.FullName}";
        _state.Text = adapter.IsAdminEnabled ? "● CONNECTED" : "● DISABLED";
        _state.ForeColor = adapter.IsAdminEnabled ? Color.FromArgb(0, 224, 164) : Color.FromArgb(242, 91, 105);
    }

    private void AttachPress(Control control)
    {
        control.MouseUp += (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                if ((DateTime.UtcNow - _lastPress).TotalMilliseconds < 150)
                {
                    return;
                }

                _lastPress = DateTime.UtcNow;
                CardPressed?.Invoke(this, EventArgs.Empty);
            }
        };
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        BorderColor = IsSelected ? _theme.Accent : _theme.Stroke;
        base.OnPaint(e);
    }
}

internal sealed class AdapterStateCard : SurfacePanel
{
    private readonly Label _kind = new() { AutoSize = true, Font = new Font("Consolas", 10F) };
    private readonly Label _name = new() { AutoSize = true, Font = new Font("Segoe UI", 18F, FontStyle.Bold) };
    private readonly Label _meta = new() { AutoSize = true, Font = new Font("Segoe UI", 11.5F) };
    private readonly Label _state = new() { AutoSize = true, Font = new Font("Segoe UI", 13F, FontStyle.Bold) };

    private UiTheme _theme = UiTheme.Dark;
    private bool _highlighted;

    public AdapterStateCard()
    {
        Padding = new Padding(16, 14, 16, 12);
        var stack = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, BackColor = Color.Transparent };
        stack.Controls.Add(_kind);
        stack.Controls.Add(new Label { Height = 14, Width = 1 });
        stack.Controls.Add(_name);
        stack.Controls.Add(_meta);
        stack.Controls.Add(new Label { Height = 12, Width = 1 });
        stack.Controls.Add(_state);
        Controls.Add(stack);
    }

    public void SetKind(string text) => _kind.Text = text;

    public override void SetTheme(UiTheme theme)
    {
        _theme = theme;
        base.SetTheme(theme);
        _kind.ForeColor = theme.Muted;
        _name.ForeColor = theme.Text;
        _meta.ForeColor = theme.Muted;
    }

    protected override float BorderThickness => _highlighted
        ? (CurrentTheme.Name == "light" ? 6f : 4f)
        : (CurrentTheme.Name == "light" ? 3f : 2f);

    public void SetPlaceholder(string title, string details)
    {
        _name.Text = title;
        _meta.Text = details;
        _state.Text = "WAITING";
        _state.ForeColor = _theme.Muted;
        _highlighted = false;
        Invalidate();
    }

    public void SetAdapter(NetworkAdapterInfo adapter, bool primary)
    {
        _name.Text = adapter.Name;
        _meta.Text = $"{adapter.Type}\n{adapter.FullName}";
        _state.Text = adapter.IsAdminEnabled ? "● CONNECTED" : "● DISABLED";
        _state.ForeColor = adapter.IsAdminEnabled ? Color.FromArgb(0, 224, 164) : Color.FromArgb(242, 91, 105);
        var activeFill = _theme.Name == "light" ? Color.FromArgb(220, 236, 248) : Color.FromArgb(31, 66, 76);
        _highlighted = adapter.IsAdminEnabled;
        FillColor = primary && adapter.IsAdminEnabled ? activeFill : _theme.Surface;
        Invalidate();
    }
}


internal sealed class ThemeMenuRenderer(UiTheme theme) : ToolStripProfessionalRenderer(new ThemeColorTable(theme))
{
}

internal sealed class ThemeColorTable(UiTheme theme) : ProfessionalColorTable
{
    public override Color ToolStripDropDownBackground => theme.Surface;
    public override Color MenuBorder => theme.Stroke;
    public override Color MenuItemSelected => Color.FromArgb(120, theme.Accent);
    public override Color MenuItemBorder => theme.Stroke;
}

internal static class NativeMethods
{
    internal const int SbVert = 1;
    internal const int WmVscroll = 0x115;
    internal const int SbThumbPosition = 4;

    [StructLayout(LayoutKind.Sequential)]
    internal struct ScrollInfo
    {
        public uint cbSize;
        public uint fMask;
        public int nMin;
        public int nMax;
        public uint nPage;
        public int nPos;
        public int nTrackPos;
    }

    [DllImport("user32.dll")]
    private static extern bool GetScrollInfo(IntPtr hWnd, int nBar, ref ScrollInfo lpScrollInfo);

    [DllImport("user32.dll")]
    internal static extern int SetScrollPos(IntPtr hWnd, int nBar, int nPos, bool bRedraw);

    [DllImport("user32.dll")]
    internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    internal static ScrollInfo GetScrollInfo(IntPtr handle)
    {
        var info = new ScrollInfo { cbSize = (uint)Marshal.SizeOf<ScrollInfo>(), fMask = 0x17 };
        _ = GetScrollInfo(handle, SbVert, ref info);
        return info;
    }
}

internal sealed class ThemedScrollBar : Control
{
    private int _value;

    public int Maximum { get; set; }
    public int ViewportSize { get; set; } = 1;
    public int Value => _value;

    private UiTheme _theme = UiTheme.Dark;

    public event Action<object, int>? ValueChanged;

    public ThemedScrollBar()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        Width = 12;
    }

    public void SetTheme(UiTheme theme)
    {
        _theme = theme;
        Invalidate();
    }

    public void SetValueClamped(int value)
    {
        var clamped = Math.Clamp(value, 0, Maximum);
        if (clamped == _value)
        {
            return;
        }

        _value = clamped;
        ValueChanged?.Invoke(this, _value);
        Invalidate();
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        var ratio = Height <= 1 ? 0 : e.Y / (double)Height;
        SetValueClamped((int)(ratio * Maximum));
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.Clear(_theme.Surface);

        if (Maximum <= 0)
        {
            return;
        }

        var thumbHeight = Math.Max(30, (int)(Height * (ViewportSize / (double)(Maximum + ViewportSize))));
        var travel = Math.Max(1, Height - thumbHeight);
        var y = (int)(travel * (_value / (double)Math.Max(1, Maximum)));

        using var brush = new SolidBrush(Color.FromArgb(210, _theme.Accent));
        using var path = SurfacePanel.RoundRect(new Rectangle(1, y + 1, Width - 3, thumbHeight - 2), 5);
        e.Graphics.FillPath(brush, path);
    }
}
