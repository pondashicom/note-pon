using System.Windows;

namespace NotePon;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        var application = new Application
        {
            ShutdownMode = ShutdownMode.OnMainWindowClose
        };

        application.Run(new MainWindow());
    }
}
