using System.Windows;
using System.Windows.Threading;

namespace NotePon;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (PowerPointWorkerHost.TryGetParentProcessId(args, out int parentProcessId))
        {
            return PowerPointWorkerHost.Run(parentProcessId);
        }

        using Mutex? singleInstance = TryCreateSingleInstanceMutex(out bool isFirstInstance);
        if (!isFirstInstance)
        {
            AppLog.Write("A second NOTE-PON instance was prevented from starting.");
            return 0;
        }

        var application = new Application
        {
            ShutdownMode = ShutdownMode.OnMainWindowClose
        };
        application.DispatcherUnhandledException += OnDispatcherUnhandledException;

        AppDomain.CurrentDomain.UnhandledException += (_, eventArgs) =>
        {
            if (eventArgs.ExceptionObject is Exception exception)
            {
                AppLog.Write("An unhandled application exception occurred.", exception);
            }
        };

        TaskScheduler.UnobservedTaskException += (_, eventArgs) =>
        {
            AppLog.Write("An unobserved task exception occurred.", eventArgs.Exception);
            eventArgs.SetObserved();
        };

        try
        {
            return application.Run(new MainWindow());
        }
        catch (Exception exception)
        {
            AppLog.Write("NOTE-PON terminated after an unrecoverable startup or dispatcher error.", exception);
            return 1;
        }
    }

    private static Mutex? TryCreateSingleInstanceMutex(out bool isFirstInstance)
    {
        try
        {
            return new Mutex(initiallyOwned: true, @"Local\NOTE-PON", out isFirstInstance);
        }
        catch (Exception exception)
        {
            AppLog.Write("The single-instance mutex could not be created. Startup will continue.", exception);
            isFirstInstance = true;
            return null;
        }
    }

    private static void OnDispatcherUnhandledException(
        object sender,
        DispatcherUnhandledExceptionEventArgs eventArgs)
    {
        AppLog.Write(
            "An unhandled UI exception occurred. NOTE-PON will shut down safely.",
            eventArgs.Exception);

        try
        {
            Application? application = Application.Current;
            if (application is null)
            {
                return;
            }

            eventArgs.Handled = true;
            application.Shutdown(1);
        }
        catch (Exception shutdownException)
        {
            eventArgs.Handled = false;
            AppLog.Write("NOTE-PON could not shut down cleanly after a UI exception.", shutdownException);
        }
    }
}
