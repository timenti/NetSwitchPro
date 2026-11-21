using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace NetworkAdapterSwitcher;

internal sealed class TrayIconFeedback : IDisposable
{
    private readonly NotifyIcon notification;
    private readonly Icon standard;
    private readonly Icon[] progress;
    private readonly Icon success;
    private readonly Icon failure;
    private readonly System.Windows.Forms.Timer timer = new();
    private bool waiting;
    private int frame;

    public TrayIconFeedback(NotifyIcon notification)
    {
        this.notification = notification;
        using var resource = typeof(TrayIconFeedback).Assembly.GetManifestResourceStream("NetworkAdapterSwitcher.TrayIcon")
            ?? throw new InvalidOperationException("The notification icon is missing.");
        standard = new Icon(resource, SystemInformation.SmallIconSize);
        progress = Enumerable.Range(0, 8).Select(index => FrameFor(standard.Size, index)).ToArray();
        success = Badge(standard.Size, true);
        failure = Badge(standard.Size, false);
        notification.Icon = standard;
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

    private static Icon FrameFor(Size size, int frame) => DrawIcon(size, (g, pen) =>
    {
        g.TranslateTransform(16, 16);
        g.RotateTransform(frame * 45);
        g.DrawArc(pen, -10, -10, 20, 20, 36, 270);
        g.DrawLines(pen, [new PointF(2, -7), new PointF(6, -8), new PointF(6.5f, -13)]);
    }, Color.White);

    private static Icon Badge(Size size, bool good) => good
        ? DrawIcon(size, (g, p) => g.DrawLines(p, [new Point(6, 16), new Point(13, 23), new Point(26, 9)]), Color.FromArgb(126, 216, 170))
        : DrawIcon(size, (g, p) =>
        {
            g.DrawLine(p, 8, 8, 24, 24);
            g.DrawLine(p, 24, 8, 8, 24);
        }, Color.FromArgb(255, 123, 123));

    private static Icon DrawIcon(Size size, Action<Graphics, Pen> paint, Color ink)
    {
        using var bitmap = new Bitmap(size.Width, size.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        using (var pen = new Pen(ink, 3.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
        {
            graphics.ScaleTransform(size.Width / 32f, size.Height / 32f);
            DesktopDrawing.Shape(graphics, new Rectangle(0, 0, 31, 31), 5, Color.FromArgb(24, 24, 24), Color.FromArgb(68, 68, 68));
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
        notification.Icon = null;
        standard.Dispose();
        foreach (var icon in progress)
            icon.Dispose();
        success.Dispose();
        failure.Dispose();
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr icon);
}
