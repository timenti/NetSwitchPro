using System.ComponentModel;

namespace NetworkAdapterSwitcher;

internal sealed class JournalView(string body) : Control
{
    private const TextFormatFlags Flags = TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
    private int offset, totalHeight;
    private bool dragging;
    private int dragOrigin, dragOffset;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    internal DesktopPalette Palette { get; set; } = DesktopPalette.Resolve(false);

    private Rectangle TextArea => Rectangle.FromLTRB(this.Px(6), this.Px(5), ClientSize.Width - this.Px(18), ClientSize.Height - this.Px(5));
    private int Limit => Math.Max(0, totalHeight - TextArea.Height);

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        AccessibleRole = AccessibleRole.StaticText;
        AccessibleDescription = body;
        TabStop = true;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Palette.Surface);
        var view = TextArea;
        if (view.Width <= 0 || view.Height <= 0)
            return;

        totalHeight = TextRenderer.MeasureText(g, body, Font, new Size(view.Width, 50000), Flags).Height;
        offset = Math.Clamp(offset, 0, Limit);
        var saved = g.Save();
        g.SetClip(view);
        var text = new Rectangle(view.Left, view.Top - offset, view.Width, totalHeight + this.Px(2));
        TextRenderer.DrawText(g, body, Font, text, Palette.Ink, Flags);
        g.Restore(saved);

        if (Limit > 0)
        {
            var track = Track();
            DesktopDrawing.Shape(g, track, this.Px(4), Palette.Field);
            DesktopDrawing.Shape(g, Thumb(track), this.Px(4), dragging ? Palette.Muted : Palette.Line);
        }
    }

    private Rectangle Track() => new(Width - this.Px(10), TextArea.Top, this.Px(7), TextArea.Height);
    private Rectangle Thumb(Rectangle track)
    {
        int height = Math.Clamp((int)((long)track.Height * track.Height / Math.Max(1, totalHeight)), Math.Min(this.Px(25), track.Height), track.Height);
        return new Rectangle(track.X, track.Y + (track.Height - height) * offset / Math.Max(1, Limit), track.Width, height);
    }

    private void ScrollTo(int value)
    {
        offset = Math.Clamp(value, 0, Limit);
        Invalidate();
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        base.OnMouseWheel(e);
        ScrollTo(offset - e.Delta * this.Px(42) / 120);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button != MouseButtons.Left)
            return;
        Focus();
        var track = Track();
        if (!track.Contains(e.Location) || Limit == 0)
            return;

        var thumb = Thumb(track);
        if (!thumb.Contains(e.Location))
            ScrollTo(offset + (e.Y > thumb.Bottom ? TextArea.Height : -TextArea.Height));
        dragging = true;
        dragOrigin = e.Y;
        dragOffset = offset;
        Capture = true;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (dragging)
        {
            var track = Track();
            int travel = Math.Max(1, track.Height - Thumb(track).Height);
            ScrollTo(dragOffset + (e.Y - dragOrigin) * Limit / travel);
        }
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        if (e.Button == MouseButtons.Left && dragging)
            Capture = false;
    }

    protected override void OnMouseCaptureChanged(EventArgs e)
    {
        base.OnMouseCaptureChanged(e);
        dragging = false;
        Invalidate();
    }

    protected override bool IsInputKey(Keys keyData) =>
        keyData is Keys.Up or Keys.Down or Keys.PageUp or Keys.PageDown or Keys.Home or Keys.End || base.IsInputKey(keyData);

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        switch (e.KeyCode)
        {
            case Keys.Up: ScrollTo(offset - this.Px(19)); break;
            case Keys.Down: ScrollTo(offset + this.Px(19)); break;
            case Keys.PageUp: ScrollTo(offset - TextArea.Height); break;
            case Keys.PageDown: ScrollTo(offset + TextArea.Height); break;
            case Keys.Home: ScrollTo(0); break;
            case Keys.End: ScrollTo(Limit); break;
            default: return;
        }
        e.Handled = true;
    }
}
