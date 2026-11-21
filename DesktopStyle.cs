using System.ComponentModel;
using System.Drawing.Drawing2D;

namespace NetworkAdapterSwitcher;

internal sealed record DesktopPalette(
    Color Window, Color Surface, Color Field, Color Hover, Color Line,
    Color Ink, Color Muted, Color Accent, Color AccentHover, Color AccentInk, Color Signal)
{
    private static readonly DesktopPalette Dark = new(
        Gray(18), Gray(25), Gray(32), Gray(42), Gray(53), Gray(246), Gray(176),
        Color.FromArgb(126, 216, 170), Color.FromArgb(152, 229, 188), Gray(18),
        Color.FromArgb(126, 216, 170));

    private static readonly DesktopPalette Light = new(
        Gray(246), Color.White, Gray(239), Gray(230), Gray(213), Gray(22), Gray(94),
        Color.FromArgb(126, 216, 170), Color.FromArgb(152, 229, 188), Gray(18),
        Color.FromArgb(25, 111, 68));

    public static DesktopPalette Resolve(bool light) => light ? Light : Dark;
    private static Color Gray(int value) => Color.FromArgb(value, value, value);
}

internal static class DesktopDrawing
{
    public static int Px(this Control control, int value) => (value * control.DeviceDpi + 48) / 96;

    public static GraphicsPath Rounded(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        int d = Math.Min(radius * 2, Math.Min(bounds.Width, bounds.Height));
        if (d <= 1)
        {
            path.AddRectangle(bounds);
            return path;
        }
        path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
        path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
        path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    public static void Shape(Graphics g, Rectangle area, int radius, Color fill, Color? stroke = null)
    {
        if (area.Width < 1 || area.Height < 1)
            return;

        var previous = g.SmoothingMode;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var path = Rounded(area, radius);
        using var brush = new SolidBrush(fill);
        g.FillPath(brush, path);
        if (stroke is { } outline)
        {
            using var pen = new Pen(outline);
            g.DrawPath(pen, path);
        }
        g.SmoothingMode = previous;
    }
}

internal class LayoutGrid : TableLayoutPanel
{
    public LayoutGrid() => DoubleBuffered = true;
}

internal sealed class SurfaceGrid : LayoutGrid
{

    protected override void OnPaintBackground(PaintEventArgs e)
    {
        e.Graphics.Clear(Parent?.BackColor ?? BackColor);
        DesktopDrawing.Shape(e.Graphics, new Rectangle(0, 0, Width - 1, Height - 1), this.Px(16), BackColor, ForeColor);
    }
}

internal enum ButtonAppearance { Quiet, Outline, Segment, Field, Primary, Confirm }

internal class DesktopButton : Button
{
    private bool hovered;
    private bool pressed;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    internal ButtonAppearance Appearance { get; init; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    internal bool Selected { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    internal string? Glyph { get; init; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    internal DesktopPalette Palette { get; set; } = DesktopPalette.Resolve(false);

    public DesktopButton()
    {
        FlatStyle = FlatStyle.Flat;
        UseVisualStyleBackColor = false;
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        Cursor = Cursors.Hand;
        Font = new Font("Segoe UI Semibold", 10);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            Font.Dispose();
        base.Dispose(disposing);
    }

    protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); hovered = true; Invalidate(); }
    protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); hovered = false; pressed = false; Invalidate(); }
    protected override void OnMouseDown(MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left && Parent is ModeSelector selector)
        {
            selector.BeginMouse(PointToScreen(e.Location));
            return;
        }
        base.OnMouseDown(e);
        pressed = e.Button == MouseButtons.Left;
        Invalidate();
    }

    protected override void WndProc(ref Message m)
    {
        if (Parent is not ModeSelector selector || !selector.HandleTouch(ref m))
            base.WndProc(ref m);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (Parent is ModeSelector selector && keyData is Keys.Left or Keys.Right)
        {
            selector.RequestMode(keyData == Keys.Left);
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
    protected override void OnMouseUp(MouseEventArgs e) { base.OnMouseUp(e); pressed = false; Invalidate(); }
    protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
    protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
    protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); Invalidate(); }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        if (Parent is ModeSelector selector)
            selector.PaintTrack(g, Left, Top);
        else
            g.Clear(Parent?.BackColor ?? Palette.Window);
        bool over = Enabled && (hovered || pressed);
        bool invert = Enabled && (Appearance is ButtonAppearance.Primary or ButtonAppearance.Confirm);
        bool field = Appearance == ButtonAppearance.Field;
        Color background = Appearance switch
        {
            ButtonAppearance.Primary => Enabled ? (over ? Palette.AccentHover : Palette.Accent) : Palette.Field,
            ButtonAppearance.Confirm => Enabled ? (over ? Color.FromArgb(225, 225, 225) : Color.White) : Palette.Field,
            ButtonAppearance.Field => over ? Palette.Hover : Palette.Field,
            ButtonAppearance.Segment => Selected || over ? Palette.Hover : Parent?.BackColor ?? Palette.Surface,
            _ => over ? Palette.Hover : Palette.Surface
        };
        Color? border = field || Selected || Appearance is ButtonAppearance.Outline or ButtonAppearance.Confirm ? Palette.Line : null;
        if (Focused && ShowFocusCues)
            border = Palette.Signal;
        var bounds = new Rectangle(1, 1, Width - 3, Height - 3);
        if (Appearance != ButtonAppearance.Segment)
            DesktopDrawing.Shape(g, bounds, this.Px(10), background, border);
        else if (Focused && ShowFocusCues)
        {
            using var focus = new Pen(Palette.Signal);
            using var outline = DesktopDrawing.Rounded(Rectangle.Inflate(bounds, -2, -2), this.Px(9));
            g.DrawPath(focus, outline);
        }

        var foreground = invert ? Palette.AccentInk :
            !Enabled || (field && !Selected) || (Appearance == ButtonAppearance.Segment && !Selected)
                ? Palette.Muted : Palette.Ink;
        var textArea = Rectangle.Inflate(bounds, -this.Px(12), 0);
        var flags = TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
        if (field)
        {
            textArea.Width -= this.Px(21);
            using var pen = new Pen(Palette.Muted, Math.Max(1f, this.Px(1))) { StartCap = LineCap.Round, EndCap = LineCap.Round };
            int x = bounds.Right - this.Px(16), y = bounds.Top + bounds.Height / 2;
            int d = this.Px(4);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.DrawLines(pen, [new Point(x - d, y - d / 2), new Point(x, y + d / 2), new Point(x + d, y - d / 2)]);
        }

        if (Glyph is { } glyph)
        {
            using var iconFont = new Font("Segoe MDL2 Assets", field ? 11 : 12, GraphicsUnit.Point);
            if (Text.Length == 0)
            {
                TextRenderer.DrawText(g, glyph, iconFont, bounds, foreground, flags | TextFormatFlags.HorizontalCenter);
                return;
            }

            int iconWidth = this.Px(20), gap = this.Px(7);
            var measureFlags = flags & ~TextFormatFlags.EndEllipsis;
            int measured = TextRenderer.MeasureText(g, Text, Font, new Size(32767, 32767), measureFlags).Width;
            int textWidth = Math.Min(measured + this.Px(3), Math.Max(0, textArea.Width - iconWidth - gap));
            int x = (Width - textWidth - iconWidth - gap) / 2;
            TextRenderer.DrawText(g, glyph, iconFont, new Rectangle(x, bounds.Top, iconWidth, bounds.Height), foreground,
                flags | TextFormatFlags.HorizontalCenter);
            textArea = new Rectangle(x + iconWidth + gap, bounds.Top, textWidth, bounds.Height);
        }

        if (textArea.Width > 0)
            TextRenderer.DrawText(g, Text, Font, textArea, foreground,
                flags | (field || Glyph is not null ? TextFormatFlags.Left : TextFormatFlags.HorizontalCenter));
    }
}

internal sealed record DesktopChoice(string Key, string Caption, string? Detail = null, bool Active = false);

internal sealed class DesktopPicker : DesktopButton
{
    private readonly PickerMenu menu = new();
    private DesktopChoice[] options = [];
    private string? selectedKey;

    public event EventHandler? SelectionChangeCommitted;
    public string? SelectedKey => selectedKey;
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    internal string Placeholder { get; set; } = "";

    public DesktopPicker()
    {
        Appearance = ButtonAppearance.Field;
        AccessibleRole = AccessibleRole.PushButton;
    }

    public void Bind(IEnumerable<DesktopChoice> choices, string? key)
    {
        options = choices.ToArray();
        SelectKey(key);
    }

    private void SelectKey(string? key)
    {
        var choice = options.FirstOrDefault(x => x.Key == key);
        selectedKey = choice?.Key;
        Selected = choice is not null;
        Text = choice?.Caption ?? Placeholder;
        AccessibleDescription = Text;
        Invalidate();
    }

    protected override void OnClick(EventArgs e)
    {
        base.OnClick(e);
        if (options.Length == 0 || menu.Visible)
            return;

        menu.Font = Font;
        menu.Padding = new Padding(this.Px(5));
        int desiredWidth = Math.Max(this.Px(165), Width);
        menu.MinimumSize = new Size(desiredWidth, 0);
        menu.Renderer = new MenuRenderer(Palette);
        menu.BackColor = Palette.Surface;
        while (menu.Items.Count > 0)
            menu.Items[0].Dispose();

        foreach (var choice in options)
        {
            var item = new PickerItem(choice, Palette, desiredWidth - menu.Padding.Horizontal, DeviceDpi, choice.Key == selectedKey);
            item.Click += (_, _) =>
            {
                SelectKey(choice.Key);
                SelectionChangeCommitted?.Invoke(this, EventArgs.Empty);
            };
            menu.Items.Add(item);
        }
        menu.Show(this, new Point(0, Height + this.Px(3)), ToolStripDropDownDirection.BelowRight);
        menu.Items.OfType<PickerItem>().FirstOrDefault(x => x.Checked)?.Select();
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData is (Keys.Alt | Keys.Down) or Keys.F4 or Keys.Down)
        {
            PerformClick();
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            menu.Dispose();
        base.Dispose(disposing);
    }

    private sealed class PickerItem : ToolStripMenuItem
    {
        private readonly DesktopChoice choice;
        private readonly DesktopPalette palette;
        private readonly Size rowSize;
        private readonly int dpi;

        internal PickerItem(DesktopChoice choice, DesktopPalette palette, int width, int dpi, bool selected)
            : base(choice.Caption.Replace("&", "&&"))
        {
            this.choice = choice;
            this.palette = palette;
            this.dpi = dpi;
            rowSize = new Size(width, Px(choice.Detail is null ? 34 : 51));
            Checked = selected;
            AccessibleName = choice.Caption;
            AccessibleDescription = choice.Detail ?? "";
            Margin = Padding.Empty;
        }

        public override Size GetPreferredSize(Size constrainingSize) => rowSize;

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            if (Selected || Checked)
                DesktopDrawing.Shape(g, new Rectangle(0, 0, Width - 1, Height - 1), Px(8),
                    Selected ? palette.Hover : palette.Field);

            const TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding |
                TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis;
            int labelWidth = Math.Max(0, Width - Px(43));
            if (choice.Detail is { } detail)
            {
                using var detailFont = new Font("Segoe UI", 9);
                TextRenderer.DrawText(g, choice.Caption, Font, new Rectangle(Px(12), Px(6), labelWidth, Px(20)), palette.Ink, flags);
                TextRenderer.DrawText(g, detail, detailFont, new Rectangle(Px(12), Px(27), labelWidth, Px(15)),
                    choice.Active ? palette.Signal : palette.Muted, flags);
            }
            else
                TextRenderer.DrawText(g, choice.Caption, Font, new Rectangle(Px(12), 0, labelWidth, Height), palette.Ink, flags);

            if (!Checked)
                return;
            var previous = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var pen = new Pen(palette.Signal, Math.Max(1, Px(2))) { StartCap = LineCap.Round, EndCap = LineCap.Round };
            int x = Width - Px(18), y = Height / 2;
            g.DrawLines(pen, [new Point(x - Px(4), y), new Point(x - Px(1), y + Px(3)), new Point(x + Px(5), y - Px(4))]);
            g.SmoothingMode = previous;
        }

        private int Px(int size) => (size * dpi + 48) / 96;
    }

    private sealed class PickerMenu : ContextMenuStrip
    {
        public PickerMenu()
        {
            ShowImageMargin = ShowCheckMargin = false;
            DropShadowEnabled = false;
        }

        protected override void OnOpening(CancelEventArgs e)
        {
            base.OnOpening(e);
            UpdateRegion();
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            UpdateRegion();
        }

        private void UpdateRegion()
        {
            if (Width < 2 || Height < 2)
                return;
            using var outline = DesktopDrawing.Rounded(ClientRectangle, this.Px(13));
            var old = Region;
            Region = new Region(outline);
            old?.Dispose();
        }
    }

    private sealed class MenuRenderer(DesktopPalette palette) : ToolStripRenderer
    {
        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
        {
            var bounds = new Rectangle(0, 0, e.ToolStrip.Width - 1, e.ToolStrip.Height - 1);
            DesktopDrawing.Shape(e.Graphics, bounds, e.ToolStrip.Px(13), palette.Surface, palette.Line);
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
        {
            e.ArrowColor = palette.Muted;
            base.OnRenderArrow(e);
        }
    }

}

internal sealed class DesktopToggle : CheckBox
{
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    internal DesktopPalette Palette { get; set; } = DesktopPalette.Resolve(false);

    public DesktopToggle()
    {
        AutoSize = false;
        Size = new Size(112, 33);
        Margin = new Padding(0, 0, 5, 0);
        Cursor = Cursors.Hand;
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnCheckedChanged(EventArgs e)
    {
        base.OnCheckedChanged(e);
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Parent?.BackColor ?? Palette.Surface);
        DesktopDrawing.Shape(g, new Rectangle(0, 0, Width - 1, Height - 1), this.Px(10), Palette.Field,
            Focused && ShowFocusCues ? Palette.Signal : Palette.Line);
        int dotsize = this.Px(10);
        var marker = new Rectangle(this.Px(10), (Height - dotsize) / 2, dotsize, dotsize);
        DesktopDrawing.Shape(g, marker, dotsize, Checked ? Palette.Signal : Palette.Field,
            Checked ? null : Palette.Muted);
        var text = new Rectangle(this.Px(26), 0, Width - this.Px(33), Height);
        TextRenderer.DrawText(g, Text, Font, text, Enabled ? Palette.Ink : Palette.Muted,
            TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
            TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
    }
}
