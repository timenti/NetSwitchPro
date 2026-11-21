using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace NetworkAdapterSwitcher;

internal sealed class TrayIconFeedback : IDisposable
{
    private readonly NotifyIcon notification;
    private readonly Icon standard;
    private readonly Icon[] progress = Enumerable.Range(0, 8).Select(FrameFor).ToArray();
    private readonly Icon success = Badge(true);
    private readonly Icon failure = Badge(false);
    private readonly System.Windows.Forms.Timer timer = new();
    private bool waiting;
    private int frame;

    public TrayIconFeedback(NotifyIcon notification, Icon standard)
    {
        this.notification = notification;
        this.standard = standard;
        timer.Tick += (_, _) =>
        {
            if (waiting)
            {
                frame = (frame + 1) % progress.Length;
                notification.Icon = progress[frame];
            }
            else
            {
                timer.Stop();
                notification.Icon = standard;
                notification.Text = "NetSwitch Pro";
            }
        };
    }

    public void Begin(string text)
    {
        timer.Stop();
        waiting = true;
        frame = 0;
        notification.Text = ShortText(text);
        notification.Icon = progress[0];
        timer.Interval = 100;
        timer.Start();
    }

    public void Complete(bool connected, string text)
    {
        timer.Stop();
        waiting = false;
        notification.Text = ShortText(text);
        notification.Icon = connected ? success : failure;
        timer.Interval = 1600;
        timer.Start();
    }

    private static string ShortText(string status) =>
        status.Length <= 120 ? status : status[..119] + "…";

    private static Icon FrameFor(int frame) => DrawIcon((g, pen) =>
    {
        g.TranslateTransform(16, 16);
        g.RotateTransform(frame * 45);
        g.DrawArc(pen, -8, -8, 16, 16, 36, 270);
        g.DrawLines(pen, [new PointF(1.5f, -6), new PointF(5, -7), new PointF(5.5f, -11)]);
    }, Color.White);

    private static Icon Badge(bool good) => good
        ? DrawIcon((g, p) => g.DrawLines(p, [new Point(8, 16), new Point(14, 22), new Point(24, 10)]), Color.FromArgb(126, 216, 170))
        : DrawIcon((g, p) =>
        {
            g.DrawLine(p, 10, 10, 22, 22);
            g.DrawLine(p, 22, 10, 10, 22);
        }, Color.FromArgb(255, 123, 123));

    private static Icon DrawIcon(Action<Graphics, Pen> paint, Color ink)
    {
        using var bitmap = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        using (var pen = new Pen(ink, 3) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
        {
            DesktopDrawing.Shape(graphics, new Rectangle(1, 1, 29, 29), 7, Color.FromArgb(24, 24, 24), Color.FromArgb(68, 68, 68));
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            paint(graphics, pen);
        }
        IntPtr handle = bitmap.GetHicon();
        try { return (Icon)Icon.FromHandle(handle).Clone(); }
        finally { _ = DestroyIcon(handle); }
    }

    public void Dispose()
    {
        timer.Dispose();
        notification.Icon = standard;
        foreach (var icon in progress)
            icon.Dispose();
        success.Dispose();
        failure.Dispose();
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr icon);
}
