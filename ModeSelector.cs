using System.ComponentModel;
using System.Runtime.InteropServices;

namespace NetworkAdapterSwitcher;

internal sealed class ModeSelector : Panel
{
    private readonly DesktopButton wifi;
    private readonly DesktopButton adapters;
    private readonly System.Windows.Forms.Timer animation = new() { Interval = 14 };
    private readonly System.Diagnostics.Stopwatch elapsed = new();
    private float position, animationFrom;
    private bool wifiSelected = true;
    private Point? down;
    private uint? touchId;

    public event Action<bool>? ModeRequested;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    internal DesktopPalette Palette { get; set; } = DesktopPalette.Resolve(false);

    public ModeSelector(DesktopButton wifi, DesktopButton adapters)
    {
        this.wifi = wifi;
        this.adapters = adapters;
        DoubleBuffered = true;
        SetStyle(ControlStyles.ResizeRedraw, true);
        wifi.Dock = adapters.Dock = DockStyle.None;
        Controls.AddRange([wifi, adapters]);
        animation.Tick += (_, _) =>
        {
            float progress = Math.Clamp((float)(elapsed.Elapsed.TotalMilliseconds / 145), 0, 1);
            float eased = 1 - MathF.Pow(1 - progress, 3);
            MoveIndicator(animationFrom + (Destination - animationFrom) * eased);
            if (progress >= 1)
                animation.Stop();
        };
    }

    private int Inset => this.Px(4);
    private float Travel => Math.Max(1, (Width - Inset * 2) / 2f);
    private float Destination => wifiSelected ? 0 : 1;

    public void SelectMode(bool showWifi, bool animate)
    {
        if (showWifi == wifiSelected && (animation.Enabled || Math.Abs(Destination - position) < .001f))
            return;
        wifiSelected = showWifi;
        Settle(animate);
    }

    private void Settle(bool animate = true)
    {
        animation.Stop();
        if (!animate || !SystemInformation.IsMenuAnimationEnabled || Math.Abs(Destination - position) < .001f)
        {
            MoveIndicator(Destination);
            return;
        }
        animationFrom = position;
        elapsed.Restart();
        animation.Start();
    }

    private void MoveIndicator(float place)
    {
        position = place;
        Invalidate(true);
    }

    protected override void OnLayout(LayoutEventArgs e)
    {
        base.OnLayout(e);
        int leftWidth = Math.Max(0, (Width - 2 * Inset) / 2);
        int height = Math.Max(0, Height - 2 * Inset);
        wifi.Bounds = new Rectangle(Inset, Inset, leftWidth, height);
        adapters.Bounds = new Rectangle(Inset + leftWidth, Inset, leftWidth, height);
    }

    internal void PaintTrack(Graphics g, int left, int top)
    {
        var saved = g.Save();
        g.TranslateTransform(-left, -top);
        using var background = new SolidBrush(Parent?.BackColor ?? Palette.Window);
        g.FillRectangle(background, ClientRectangle);
        DesktopDrawing.Shape(g, new Rectangle(0, 0, Width - 1, Height - 1), this.Px(16), Palette.Surface, Palette.Line);
        var indicator = new Rectangle(Inset + (int)MathF.Round(position * Travel), Inset,
            (int)Travel - 1, Height - Inset * 2 - 1);
        DesktopDrawing.Shape(g, indicator, this.Px(11), Palette.Hover, Palette.Line);
        g.Restore(saved);
    }

    protected override void OnPaintBackground(PaintEventArgs e) => PaintTrack(e.Graphics, 0, 0);

    internal void BeginMouse(Point screen)
    {
        if (!Enabled || touchId.HasValue)
            return;
        BeginGesture(screen);
        Capture = true;
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button == MouseButtons.Left)
            BeginMouse(PointToScreen(e.Location));
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (Capture)
            MoveGesture(PointToScreen(e.Location));
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        if (!Capture || e.Button != MouseButtons.Left)
            return;
        FinishGesture(PointToScreen(e.Location));
        Capture = false;
    }

    protected override void OnMouseCaptureChanged(EventArgs e)
    {
        base.OnMouseCaptureChanged(e);
        if (!Capture && down.HasValue && !touchId.HasValue)
            CancelGesture();
    }

    private void BeginGesture(Point screen)
    {
        animation.Stop();
        down = PointToClient(screen);
    }

    private void MoveGesture(Point screen)
    {
        if (down is not { } start)
            return;
        var current = PointToClient(screen);
        MoveIndicator(Math.Clamp(Destination + (current.X - start.X) / Travel, 0, 1));
    }

    private void FinishGesture(Point screen)
    {
        if (down is not { } start)
            return;
        down = null;
        var current = PointToClient(screen);
        int dx = current.X - start.X, dy = current.Y - start.Y;
        int threshold = this.Px(16);
        if (Math.Abs(dx) >= threshold && Math.Abs(dx) >= Math.Abs(dy))
            RequestMode(dx < 0);
        else if (Math.Abs(dx) < threshold && Math.Abs(dy) < threshold && ClientRectangle.Contains(current))
            RequestMode(current.X < Width / 2);
        Settle();
    }

    internal void RequestMode(bool showWifi)
    {
        if (Enabled)
            ModeRequested?.Invoke(showWifi);
    }

    private void CancelGesture()
    {
        down = null;
        touchId = null;
        Settle();
    }

    protected override void WndProc(ref Message m)
    {
        if (!HandleTouch(ref m))
            base.WndProc(ref m);
    }

    internal bool HandleTouch(ref Message message)
    {
        const int update = 0x0245, press = 0x0246, release = 0x0247, captureChanged = 0x024C;
        if (message.Msg is not (update or press or release or captureChanged))
            return false;
        uint id = (uint)(message.WParam.ToInt64() & 0xffff);
        if (message.Msg == press && (!Enabled || touchId.HasValue || !GetPointerType(id, out uint type) || type != 2))
            return false;
        if (message.Msg != press && id != touchId)
            return false;

        long value = message.LParam.ToInt64();
        var screen = new Point(unchecked((short)value), unchecked((short)(value >> 16)));
        switch (message.Msg)
        {
            case press:
                touchId = id;
                BeginGesture(screen);
                break;
            case update:
                MoveGesture(screen);
                break;
            case release:
                touchId = null;
                FinishGesture(screen);
                break;
            case captureChanged:
                CancelGesture();
                break;
        }
        message.Result = IntPtr.Zero;
        return true;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            animation.Dispose();
        base.Dispose(disposing);
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetPointerType(uint pointerId, out uint type);
}
