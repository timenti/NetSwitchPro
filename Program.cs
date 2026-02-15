using System.Globalization;

namespace NetworkAdapterSwitcher;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("ru-RU");
        Application.Run(new MainForm());
    }
}
