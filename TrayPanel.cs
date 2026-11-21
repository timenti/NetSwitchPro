namespace NetworkAdapterSwitcher;

internal sealed record TrayState(
    string Language, bool IsLight, string Connection, bool Connected, string[] Pair, string[] Saved,
    bool Busy, string? Feedback, string? Target);

internal sealed class TrayPanel : Form
{
    private const int PageSize = 5;

    private readonly Label caption = Label(9);
    private readonly Label connection = Label(16, true);
    private readonly Label section = Label(10, true);
    private readonly Label feedback = Label(9);
    private readonly Label pageNumber = Label(9);
    private readonly DesktopButton[] profiles = Enumerable.Range(0, PageSize).Select(_ => Button(ButtonAppearance.Outline)).ToArray();
    private readonly DesktopButton switchPair = Button(ButtonAppearance.Primary);
    private readonly DesktopButton all = Button();
    private readonly DesktopButton previous = Button();
    private readonly DesktopButton next = Button();
    private readonly DesktopButton open = Button();
    private readonly DesktopButton exit = Button();
    private readonly ToolTip tips = new();
    private TrayState state = new("English", false, "", false, [], [], true, null, null);
    private DesktopPalette palette = DesktopPalette.Resolve(false);
    private Point anchor;
    private bool showAll;
    private int page;
    private bool positioning;

    public event Action<string>? ProfileRequested;
    public event Action? OpenRequested;
    public event Action? ExitRequested;

    public TrayPanel()
    {
        Text = "NetSwitch Pro";
        AutoScaleMode = AutoScaleMode.None;
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        TopMost = true;
        DoubleBuffered = true;

        Controls.AddRange([caption, connection, switchPair, section, .. profiles, all, previous, pageNumber, next, feedback, open, exit]);
        switchPair.Click += (_, _) =>
        {
            if (!state.Busy && state.Target is { } target)
                ProfileRequested?.Invoke(target);
        };
        foreach (var row in profiles)
            row.Click += (_, _) =>
            {
                if (row.Tag is string target && !state.Busy)
                    ProfileRequested?.Invoke(target);
            };
        all.Click += (_, _) =>
        {
            showAll = !showAll;
            page = 0;
            Render();
        };
        previous.Click += (_, _) => ChangePage(-1);
        next.Click += (_, _) => ChangePage(1);
        open.Click += (_, _) => OpenRequested?.Invoke();
        exit.Click += (_, _) => ExitRequested?.Invoke();
        Deactivate += (_, _) => Hide();
        DpiChanged += (_, _) => Render();
    }

    public void OpenAt(Point screen, TrayState nextState)
    {
        anchor = screen;
        showAll = false;
        page = 0;
        // Create the HWND on the monitor with the tray before measuring its mixed-DPI layout.
        Location = Screen.FromPoint(screen).WorkingArea.Location;
        UpdateState(nextState);
        Location = PlaceNear(Screen.FromPoint(anchor).WorkingArea, anchor, Size, this.Px(10));
        Show();
        Render();
        Activate();
    }

    public void UpdateState(TrayState nextState)
    {
        state = nextState;
        palette = DesktopPalette.Resolve(state.IsLight);
        Render();
    }

    private string W(Word word) => UiWords.Get(state.Language, word);

    private void Render()
    {
        if (positioning || IsDisposed)
            return;
        positioning = true;
        SuspendLayout();
        try
        {
            BackColor = palette.Surface;
            ForeColor = palette.Ink;
            foreach (Control control in Controls)
            {
                control.BackColor = palette.Surface;
                control.ForeColor = palette.Ink;
                if (control is DesktopButton button)
                    button.Palette = palette;
            }

            bool expanded = showAll || state.Pair.Length < 2;
            var available = expanded ? state.Saved : state.Pair;
            int pageCount = Math.Max(1, (available.Length + PageSize - 1) / PageSize);
            page = Math.Clamp(page, 0, pageCount - 1);
            var slice = available.Skip(page * PageSize).Take(PageSize).ToArray();

            int width = this.Px(374), pad = this.Px(16), inner = width - 2 * pad;
            int y = this.Px(15);
            caption.ForeColor = palette.Muted;
            caption.Text = W(Word.Active);
            Place(caption, pad, y, inner, 20);
            y += this.Px(20);

            connection.Text = (state.Connected ? "●  " : "") + state.Connection;
            connection.ForeColor = state.Connected ? palette.Signal : palette.Ink;
            Place(connection, pad, y, inner, 33);
            tips.SetToolTip(connection, state.Connection);
            y += this.Px(41);

            switchPair.Text = state.Target is { } targetName ? W(Word.Switch) + " " + targetName : W(Word.SwitchPair);
            tips.SetToolTip(switchPair, state.Target is null ? W(Word.PickWifi) : switchPair.Text);
            switchPair.Enabled = state.Target is not null && !state.Busy;
            Place(switchPair, pad, y, inner, 41);
            y += this.Px(51);

            section.Text = expanded ? W(Word.SavedNetworks) : W(Word.PairNetworks);
            Place(section, pad, y, inner, 23);
            y += this.Px(26);

            for (int i = 0; i < profiles.Length; i++)
            {
                var row = profiles[i];
                bool exists = i < slice.Length;
                row.Visible = exists;
                if (!exists)
                    continue;
                string target = slice[i];
                bool active = state.Connected && target == state.Connection;
                row.Tag = target;
                row.Text = active ? target + "  ·  " + W(Word.Connected) : target;
                row.Selected = active;
                row.Enabled = !active && !state.Busy;
                tips.SetToolTip(row, row.Text);
                Place(row, pad, y, inner, 37);
                y += this.Px(41);
            }

            if (state.Saved.Length == 0)
            {
                feedback.Text = state.Feedback ?? W(Word.NoSavedNetworks);
                feedback.Visible = true;
                feedback.ForeColor = palette.Muted;
                Place(feedback, pad, y, inner, 39);
                y += this.Px(43);
            }
            else
            {
                bool showFeedback = state.Feedback is not null;
                feedback.Visible = showFeedback;
                if (showFeedback)
                {
                    feedback.Text = state.Feedback;
                    feedback.ForeColor = palette.Muted;
                    Place(feedback, pad, y, inner, 37);
                    tips.SetToolTip(feedback, feedback.Text);
                    y += this.Px(39);
                }
            }

            bool paginate = pageCount > 1;
            previous.Visible = next.Visible = pageNumber.Visible = paginate;
            if (paginate)
            {
                int step = this.Px(41);
                previous.Text = "‹";
                next.Text = "›";
                previous.AccessibleName = W(Word.Previous);
                next.AccessibleName = W(Word.Next);
                previous.Enabled = page > 0;
                next.Enabled = page + 1 < pageCount;
                pageNumber.Text = $"{page + 1} / {pageCount}";
                pageNumber.TextAlign = ContentAlignment.MiddleCenter;
                pageNumber.ForeColor = palette.Muted;
                Place(previous, pad, y, step, 31);
                Place(pageNumber, pad + step, y, inner - 2 * step, 31);
                Place(next, width - pad - step, y, step, 31);
                y += this.Px(35);
            }

            bool hasMore = state.Pair.Length == 2 && state.Saved.Length > state.Pair.Length;
            all.Visible = hasMore;
            if (hasMore)
            {
                all.Text = showAll ? W(Word.PairNetworks) : W(Word.AllSavedNetworks);
                Place(all, pad, y, inner, 31);
                y += this.Px(35);
            }

            open.Text = W(Word.OpenApp);
            exit.Text = W(Word.Exit);
            exit.Enabled = !state.Busy;
            int exitWidth = this.Px(88), gap = this.Px(7);
            Place(open, pad, y, inner - exitWidth - gap, 34);
            Place(exit, width - pad - exitWidth, y, exitWidth, 34);
            ClientSize = new Size(width, y + this.Px(48));
            if (Visible)
                Location = PlaceNear(Screen.FromPoint(anchor).WorkingArea, anchor, Size, this.Px(10));
        }
        finally
        {
            ResumeLayout();
            positioning = false;
            Invalidate(true);
        }
    }

    private void Place(Control control, int x, int y, int width, int height) =>
        control.SetBounds(x, y, Math.Max(1, width), this.Px(height));

    internal static Point PlaceNear(Rectangle area, Point anchor, Size size, int gap)
    {
        int top = anchor.Y - gap - size.Height;
        if (top < area.Top)
            top = anchor.Y + gap;
        return new Point(
            Math.Clamp(anchor.X + gap - size.Width, area.Left, Math.Max(area.Left, area.Right - size.Width)),
            Math.Clamp(top, area.Top, Math.Max(area.Top, area.Bottom - size.Height)));
    }

    private void ChangePage(int step)
    {
        if (!showAll && state.Pair.Length == 2)
            return;
        int max = Math.Max(0, (state.Saved.Length - 1) / PageSize);
        int index = Math.Clamp(page + step, 0, max);
        if (index != page)
        {
            page = index;
            Render();
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Escape: Hide(); return true;
            case Keys.PageUp: ChangePage(-1); return true;
            case Keys.PageDown: ChangePage(1); return true;
            default: return base.ProcessCmdKey(ref msg, keyData);
        }
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        base.OnMouseWheel(e);
        ChangePage(e.Delta > 0 ? -1 : 1);
    }

    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        if (ClientSize.Width < 2 || ClientSize.Height < 2)
            return;
        using var shape = DesktopDrawing.Rounded(ClientRectangle, this.Px(15));
        Region = new Region(shape);
    }

    protected override void OnPaintBackground(PaintEventArgs e)
    {
        e.Graphics.Clear(palette.Surface);
        DesktopDrawing.Shape(e.Graphics, new Rectangle(0, 0, Width - 1, Height - 1), this.Px(15), palette.Surface, palette.Line);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            tips.Dispose();
        base.Dispose(disposing);
    }

    private static Label Label(float size, bool bold = false) => new()
    {
        AutoEllipsis = true,
        Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular),
        TextAlign = ContentAlignment.MiddleLeft
    };

    private static DesktopButton Button(ButtonAppearance appearance = ButtonAppearance.Quiet) => new()
    {
        Appearance = appearance,
        TextAlign = appearance == ButtonAppearance.Outline ? ContentAlignment.MiddleLeft : ContentAlignment.MiddleCenter
    };
}
